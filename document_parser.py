from __future__ import annotations

from io import BytesIO
from typing import Callable

import fitz
import numpy as np
from docx import Document
from PIL import Image


def ocr_pdf(content: bytes, ocr_reader, max_pages: int = 20) -> str:
    lines: list[str] = []
    with fitz.open(stream=content, filetype="pdf") as document:
        for page in list(document)[:max_pages]:
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = Image.open(BytesIO(pixmap.tobytes("png")))
            detected = ocr_reader.readtext(np.asarray(image), detail=0, paragraph=True)
            lines.extend(str(item) for item in detected)
    return "\n".join(lines)


def extract_document_text(
    uploaded_file,
    get_ocr_reader: Callable | None = None,
    minimum_native_characters: int = 80,
) -> tuple[str, bool]:
    extension = uploaded_file.name.lower().rsplit(".", 1)[-1]
    content = uploaded_file.getvalue()
    if extension == "pdf":
        with fitz.open(stream=content, filetype="pdf") as document:
            # Visual reading order is essential for multi-column/Canva résumés.
            # Without sort=True, headings and bullets can be extracted out of order,
            # which produces misleading scoring and badly structured tailored files.
            text = "\n".join(page.get_text(sort=True) for page in document)
        if len(text.strip()) >= minimum_native_characters or get_ocr_reader is None:
            return text, False
        return ocr_pdf(content, get_ocr_reader()), True
    if extension == "docx":
        document = Document(BytesIO(content))
        return "\n".join(paragraph.text for paragraph in document.paragraphs), False
    if extension == "txt":
        return content.decode("utf-8", errors="replace"), False
    raise ValueError(f"Unsupported document: {uploaded_file.name}")
