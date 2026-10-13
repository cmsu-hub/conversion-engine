import os
import io
import re
import json
import uuid
import shutil
import zipfile
import traceback
import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from PIL import Image
import pdf2docx
import pdfplumber
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Bi-Directional Pure Python Conversion Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
@app.get("/health")
@app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "service": "conversion-engine",
        "engine": "python",
        "supported_conversions": {
            "pdf": ["docx", "doc", "xlsx", "xls", "pptx", "ppt"]
        },
        "supported_compressions": ["pdf", "jpg", "jpeg", "png", "webp", "docx", "pptx", "xlsx"]
    }

UPLOAD_DIR = "./uploads"
OUTPUT_DIR = "./outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# CONVERSION ROUTINES
# -------------------------------------------------------------

def pdf_to_docx(input_path: str, output_path: str):
    cv = pdf2docx.Converter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()
    return output_path

def pdf_to_excel(input_path: str, output_path: str):
    all_tables = []
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                })
            
            for t_idx, table in enumerate(tables):
                if not table or len(table) < 1:
                    continue
                
                cleaned_rows = []
                for row in table:
                    if row and any(cell is not None and str(cell).strip() != "" for cell in row):
                        cleaned_rows.append([str(cell).strip() if cell is not None else "" for cell in row])
                
                if not cleaned_rows:
                    continue
                
                first_row = cleaned_rows[0]
                num_cols = len(first_row)
                
                headers = []
                seen_headers = {}
                for idx in range(num_cols):
                    raw_val = first_row[idx] if idx < len(first_row) else ""
                    val = raw_val if raw_val != "" else f"Column_{idx + 1}"
                    if val in seen_headers:
                        seen_headers[val] += 1
                        val = f"{val}_{seen_headers[val]}"
                    else:
                        seen_headers[val] = 1
                    headers.append(val)
                
                data_rows = cleaned_rows[1:] if len(cleaned_rows) > 1 else []
                aligned_data = []
                for row in data_rows:
                    if len(row) < num_cols:
                        row = row + [""] * (num_cols - len(row))
                    elif len(row) > num_cols:
                        row = row[:num_cols]
                    aligned_data.append(row)
                
                df = pd.DataFrame(aligned_data, columns=headers)
                all_tables.append((f"Page_{page_num+1}_T{t_idx+1}", df))

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if all_tables:
            for idx, (sheet_title, df) in enumerate(all_tables):
                safe_title = sheet_title[:31]
                df.to_excel(writer, sheet_name=safe_title, index=False)
        else:
            raw_lines = []
            with pdfplumber.open(input_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            if line.strip():
                                raw_lines.append([line.strip()])
            if raw_lines:
                df = pd.DataFrame(raw_lines, columns=["Extracted Text"])
                df.to_excel(writer, sheet_name="Extracted Data", index=False)
            else:
                pd.DataFrame({"Note": ["No extractable data found in PDF"]}).to_excel(writer, sheet_name="Sheet1", index=False)

    return output_path

def pdf_to_pptx(input_path: str, output_path: str):
    doc = fitz.open(input_path)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]

    temp_files = []
    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=150)
            temp_img_path = os.path.join(UPLOAD_DIR, f"temp_{uuid.uuid4().hex}_{page_num}.png")
            pix.save(temp_img_path)
            temp_files.append(temp_img_path)

            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(temp_img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

        prs.save(output_path)
    finally:
        doc.close()
        for f in temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass
    return output_path

# -------------------------------------------------------------
# UNIVERSAL DOCUMENT RESIZING & COMPRESSION ROUTINES
# -------------------------------------------------------------

def compress_pdf(input_path: str, output_path: str, quality: int = 70, level: str = "medium", max_width: int = 1600):
    """
    Compresses PDF documents.
    Preserves searchable vector text while downsampling embedded images and deflating streams.
    """
    doc = fitz.open(input_path)
    
    if level == "aggressive" or level == "high":
        # Full raster page downsampling for maximum file size reduction
        opt_doc = fitz.open()
        for page in doc:
            pix = page.get_pixmap(dpi=110)
            img_bytes = pix.tobytes("jpeg", quality=quality)
            new_page = opt_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img_bytes)
        opt_doc.save(output_path, garbage=4, deflate=True)
        doc.close()
        opt_doc.close()
        return output_path

    # Standard / Medium: Downsample embedded images without degrading text vector sharpness
    for page in doc:
        image_list = page.get_images(full=True)
        for img_info in image_list:
            xref = img_info[0]
            try:
                base_image = doc.extract_image(xref)
                if base_image:
                    image_bytes = base_image["image"]
                    pil_img = Image.open(io.BytesIO(image_bytes))
                    if pil_img.width > max_width or pil_img.height > max_width or len(image_bytes) > 80_000:
                        pil_img = pil_img.convert("RGB")
                        if pil_img.width > max_width:
                            ratio = max_width / float(pil_img.width)
                            new_h = int(float(pil_img.height) * float(ratio))
                            pil_img = pil_img.resize((max_width, new_h), Image.Resampling.LANCZOS)
                        out_io = io.BytesIO()
                        pil_img.save(out_io, format="JPEG", quality=quality, optimize=True)
                        new_bytes = out_io.getvalue()
                        if len(new_bytes) < len(image_bytes):
                            doc.update_stream(xref, new_bytes)
            except Exception:
                pass

    doc.save(output_path, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
    doc.close()
    return output_path

def resize_image(input_path: str, output_path: str, max_width: int = 1920, quality: int = 75):
    """
    Resizes image dimensions and compresses quality using Lanczos resampling.
    """
    try:
        with Image.open(input_path) as img:
            img = img.convert("RGB")
            if img.width > max_width:
                ratio = max_width / float(img.width)
                new_height = int(float(img.height) * float(ratio))
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            img.save(output_path, "JPEG", optimize=True, quality=quality)
        return output_path
    except Exception:
        # If PIL cannot parse stream or format, compress raw bytes or preserve
        try:
            with open(input_path, 'rb') as f_in:
                raw_bytes = f_in.read()
            import zlib
            compressed = zlib.compress(raw_bytes, level=9)
            if len(compressed) < len(raw_bytes):
                with open(output_path, 'wb') as f_out:
                    f_out.write(compressed)
                return output_path
        except Exception:
            pass
        shutil.copyfile(input_path, output_path)
        return output_path

def compress_office_file(input_path: str, output_path: str, max_width: int = 1600, quality: int = 70):
    """
    Compresses Microsoft Office documents (.docx, .pptx, .xlsx) by:
    1. Compressing and downsampling embedded media images
    2. Minifying XML content trees
    3. Repacking archive with maximum ZIP deflate compression
    """
    if not zipfile.is_zipfile(input_path):
        # Non-zip container: deflate raw stream or copy directly
        try:
            with open(input_path, 'rb') as f_in:
                raw_bytes = f_in.read()
            import zlib
            compressed = zlib.compress(raw_bytes, level=9)
            if len(compressed) < len(raw_bytes):
                with open(output_path, 'wb') as f_out:
                    f_out.write(compressed)
                return output_path
        except Exception:
            pass
        shutil.copyfile(input_path, output_path)
        return output_path

    with zipfile.ZipFile(input_path, 'r') as in_zip, \
         zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as out_zip:
        for item in in_zip.infolist():
            data = in_zip.read(item.filename)
            ext = os.path.splitext(item.filename)[1].lower()
            if ext in ['.jpg', '.jpeg', '.png', '.webp', '.bmp']:
                try:
                    img = Image.open(io.BytesIO(data))
                    orig_len = len(data)
                    img = img.convert("RGB")
                    if img.width > max_width:
                        ratio = max_width / float(img.width)
                        new_height = int(float(img.height) * float(ratio))
                        img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG', optimize=True, quality=quality)
                    compressed_data = img_byte_arr.getvalue()
                    if len(compressed_data) < orig_len:
                        data = compressed_data
                except Exception:
                    pass
            elif ext in ['.xml', '.rels']:
                try:
                    text = data.decode('utf-8')
                    minified = re.sub(r'>\s+<', '><', text).strip()
                    data = minified.encode('utf-8')
                except Exception:
                    pass
            out_zip.writestr(item, data)

    if os.path.exists(output_path) and os.path.getsize(output_path) > os.path.getsize(input_path):
        shutil.copyfile(input_path, output_path)
    return output_path

def compress_text_file(input_path: str, output_path: str):
    """
    Minifies text, JSON, HTML, CSV, and XML document files.
    """
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.json':
        try:
            obj = json.loads(content)
            minified = json.dumps(obj, separators=(',', ':'))
        except Exception:
            minified = content
    elif ext in ['.html', '.htm', '.xml']:
        minified = re.sub(r'<!--[\s\S]*?-->', '', content)
        minified = re.sub(r'>\s+<', '><', minified).strip()
    elif ext == '.csv':
        lines = [line.strip() for line in content.splitlines() if line.strip()]
        minified = '\n'.join(lines)
    else:
        minified = re.sub(r'[ \t]+', ' ', content)
        minified = re.sub(r'\n{3,}', '\n\n', minified).strip()
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(minified)
    return output_path

# -------------------------------------------------------------
# API ROUTES
# -------------------------------------------------------------

def cleanup_files(*paths):
    for p in paths:
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass

@app.post("/api/convert")
async def convert_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...), 
    target_format: str = Form(...)
):
    raw_name = os.path.basename(file.filename or "document")
    base_name, ext = os.path.splitext(raw_name)
    ext = ext.lower()
    target_format = target_format.strip(". ").lower()

    unique_prefix = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_DIR, f"{unique_prefix}_{raw_name}")
    out_filename = f"{base_name}.{target_format}"
    output_path = os.path.join(OUTPUT_DIR, f"{unique_prefix}_{out_filename}")

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if ext == ".pdf" and target_format in ["doc", "docx"]:
            pdf_to_docx(input_path, output_path)
        elif ext == ".pdf" and target_format in ["xls", "xlsx"]:
            pdf_to_excel(input_path, output_path)
        elif ext == ".pdf" and target_format in ["ppt", "pptx"]:
            pdf_to_pptx(input_path, output_path)
        else:
            raise HTTPException(status_code=400, detail=f"Conversion from {ext} to {target_format} not supported.")

        background_tasks.add_task(cleanup_files, input_path, output_path)
        return FileResponse(output_path, filename=out_filename)
    except HTTPException as he:
        cleanup_files(input_path, output_path)
        raise he
    except Exception as e:
        cleanup_files(input_path, output_path)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/compress")
@app.post("/api/resize")
async def compress_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...), 
    max_width: int = Form(1920),
    quality: int = Form(75),
    level: str = Form("medium")
):
    raw_name = os.path.basename(file.filename or "document")
    base_name, ext = os.path.splitext(raw_name)
    ext = ext.lower()

    unique_prefix = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_DIR, f"{unique_prefix}_{raw_name}")
    out_filename = f"compressed_{base_name}.jpg" if ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"] else f"compressed_{raw_name}"
    output_path = os.path.join(OUTPUT_DIR, f"{unique_prefix}_{out_filename}")

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if ext == ".pdf":
            compress_pdf(input_path, output_path, quality=quality, level=level, max_width=max_width)
        elif ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"]:
            resize_image(input_path, output_path, max_width=max_width, quality=quality)
        elif ext in [".pptx", ".docx", ".xlsx", ".pptm", ".docm", ".xlsm"]:
            compress_office_file(input_path, output_path, max_width=max_width, quality=quality)
        elif ext in [".json", ".csv", ".xml", ".txt", ".html", ".htm", ".md", ".sql"]:
            compress_text_file(input_path, output_path)
        else:
            raise HTTPException(status_code=400, detail=f"Compression not supported for {ext} files.")

        background_tasks.add_task(cleanup_files, input_path, output_path)
        return FileResponse(output_path, filename=out_filename)
    except HTTPException as he:
        cleanup_files(input_path, output_path)
        raise he
    except Exception as e:
        cleanup_files(input_path, output_path)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)