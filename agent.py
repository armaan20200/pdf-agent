"""
ReAct-style AI Agent for PDF operations
Handles all iLovePDF-style tools with real Python backends
"""

import os
import re
import uuid
from pathlib import Path
from typing import Optional
import json

from pdf_tools import PDFTools


class PDFAgent:
    def __init__(self, temp_dir: Path, active_files: dict, result_files: dict):
        self.temp_dir = temp_dir
        self.active_files = active_files
        self.result_files = result_files
        self.tools = PDFTools(temp_dir)
        self._openai_available = False
        self._setup_openai()

    def _setup_openai(self):
        try:
            import openai
            api_key = os.environ.get("OPENAI_API_KEY") or os.environ.get("REPLIT_AI_API_KEY")
            if api_key:
                self.client = openai.OpenAI(
                    api_key=api_key,
                    base_url=os.environ.get("OPENAI_API_BASE", "https://api.openai.com/v1"),
                )
                self._openai_available = True
        except Exception:
            pass

    def _get_files_by_ids(self, file_ids: list[str]) -> list[dict]:
        return [self.active_files[fid] for fid in file_ids if fid in self.active_files]

    def _get_all_files(self) -> list[dict]:
        return list(self.active_files.values())

    def _parse_page_range(self, text: str) -> tuple[int, int] | None:
        patterns = [
            r'pages?\s+(\d+)\s*[-–to]+\s*(\d+)',
            r'pages?\s+(\d+)\s+(?:through|thru|to)\s+(\d+)',
            r'page\s+(\d+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if len(match.groups()) == 2:
                    return int(match.group(1)), int(match.group(2))
                else:
                    page = int(match.group(1))
                    return page, page
        return None

    def _parse_page_list(self, text: str, preserve_order: bool = False) -> list[int]:
        """Parse '1, 3, 5-7' into [1, 3, 5, 6, 7]. preserve_order keeps user's ordering."""
        pages = []
        for part in re.findall(r'(\d+)\s*[-–]\s*(\d+)|(\d+)', text):
            if part[0] and part[1]:
                pages.extend(range(int(part[0]), int(part[1]) + 1))
            elif part[2]:
                pages.append(int(part[2]))
        if preserve_order:
            seen: set[int] = set()
            result: list[int] = []
            for p in pages:
                if p not in seen:
                    seen.add(p)
                    result.append(p)
            return result
        return sorted(set(pages))

    def _store_result(self, path: str, name: str) -> dict:
        file_id = str(uuid.uuid4())
        self.result_files[file_id] = {"id": file_id, "path": path, "name": name}
        return {"id": file_id, "name": name}

    def _get_input_files(self, file_ids: list[str]) -> list[dict]:
        if file_ids:
            files = self._get_files_by_ids(file_ids)
        else:
            files = self._get_all_files()
        return files

    def _first_pdf(self, file_ids: list[str]) -> dict | None:
        files = self._get_input_files(file_ids)
        pdfs = [f for f in files if f["path"].lower().endswith(".pdf")]
        return pdfs[0] if pdfs else None

    def _all_pdfs(self, file_ids: list[str]) -> list[dict]:
        files = self._get_input_files(file_ids)
        return [f for f in files if f["path"].lower().endswith(".pdf")]

    # ─────────────────────────────────────────────
    # INTENT DETECTION
    # ─────────────────────────────────────────────

    def _detect_intent(self, message: str) -> str:
        msg = message.lower()

        # Context messages from tool selection (hidden, exact matches)
        if "active tool: merge pdf" in msg:
            return "context_merge"
        if "active tool: split pdf" in msg:
            return "context_split"
        if "active tool: remove pages" in msg:
            return "context_remove_pages"
        if "active tool: extract pages" in msg:
            return "context_extract"
        if "active tool: organize pages" in msg:
            return "context_organize"
        if "active tool: rotate pdf" in msg:
            return "context_rotate"
        if "active tool: watermark" in msg:
            return "context_watermark"
        if "active tool: page numbers" in msg:
            return "context_page_numbers"
        if "active tool: jpg to pdf" in msg or "active tool: image to pdf" in msg:
            return "context_jpg_to_pdf"
        if "active tool: word to pdf" in msg:
            return "context_word_to_pdf"
        if "active tool: powerpoint to pdf" in msg or "active tool: pptx to pdf" in msg:
            return "context_pptx_to_pdf"
        if "active tool: excel to pdf" in msg:
            return "context_excel_to_pdf"
        if "active tool: html to pdf" in msg:
            return "context_html_to_pdf"
        if "active tool: pdf to jpg" in msg:
            return "context_pdf_to_jpg"
        if "active tool: pdf to word" in msg:
            return "context_pdf_to_word"
        if "active tool: pdf to powerpoint" in msg:
            return "context_pdf_to_pptx"
        if "active tool: pdf to excel" in msg:
            return "context_pdf_to_excel"
        if "active tool: pdf/a" in msg or "active tool: pdf to pdf/a" in msg:
            return "context_pdf_to_pdfa"
        if "active tool: compress pdf" in msg:
            return "context_compress"
        if "active tool: repair pdf" in msg:
            return "context_repair"
        if "active tool: sign pdf" in msg:
            return "context_sign"
        if "active tool: unlock pdf" in msg:
            return "context_unlock"
        if "active tool: protect pdf" in msg:
            return "context_protect"
        if "active tool: ai summarizer" in msg:
            return "context_summarize"
        if "active tool: ai translate" in msg:
            return "context_translate"

        # ORGANIZE
        if re.search(r'\b(merge|combine|join)\b', msg):
            return "merge"
        if re.search(r'\bsplit\b.*\bpdf\b|\bpdf\b.*\bsplit\b|\bseparate\s+pages\b', msg):
            return "split"
        if re.search(r'\b(remove|delete)\s+(pages?|page)\b', msg):
            return "remove_pages"
        if re.search(r'\b(extract|pull out|get)\s+(pages?|page)\b', msg):
            return "extract"
        if re.search(r'\b(reorder|reorganize|organize)\s+pages?\b', msg):
            return "organize"
        if re.search(r'\brotate\b', msg):
            return "rotate"
        if re.search(r'\bwatermark\b', msg):
            return "watermark"
        if re.search(r'\bpage\s+number(s|ing)?\b|\badd\s+number(s)?\b', msg):
            return "page_numbers"

        # CONVERT TO PDF
        if re.search(r'\bjpg?\s+to\s+pdf\b|\bimage\s+to\s+pdf\b|\bpng\s+to\s+pdf\b|convert.*(jpg|jpeg|image|png).*to.*pdf', msg):
            return "jpg_to_pdf"
        if re.search(r'\bword\s+to\s+pdf\b|docx?\s+to\s+pdf\b|convert.*\.(docx?)', msg):
            return "word_to_pdf"
        if re.search(r'\b(ppt|pptx|powerpoint)\s+to\s+pdf\b|convert.*(pptx?|powerpoint).*pdf', msg):
            return "pptx_to_pdf"
        if re.search(r'\bexcel\s+to\s+pdf\b|\bxlsx?\s+to\s+pdf\b|convert.*\.(xlsx?)', msg):
            return "excel_to_pdf"
        if re.search(r'\bhtml\s+to\s+pdf\b|\burl\s+to\s+pdf\b|convert.*html', msg):
            return "html_to_pdf"

        # CONVERT FROM PDF
        if re.search(r'\bpdf\s+to\s+(jpg|jpeg|image|png|picture)\b', msg):
            return "pdf_to_jpg"
        if re.search(r'\bpdf\s+to\s+word\b|\bpdf\s+to\s+docx?\b', msg):
            return "pdf_to_word"
        if re.search(r'\bpdf\s+to\s+(ppt|pptx|powerpoint)\b', msg):
            return "pdf_to_pptx"
        if re.search(r'\bpdf\s+to\s+excel\b|\bpdf\s+to\s+xlsx?\b', msg):
            return "pdf_to_excel"
        if re.search(r'\bpdf/a\b|\barchiv', msg):
            return "pdf_to_pdfa"
        if re.search(r'\bextract\s+text\b|\bpdf\s+to\s+text\b|\bconvert.*text\b', msg):
            return "pdf_to_text"

        # OPTIMIZE
        if re.search(r'\bcompress\b|\breduce\s+(size|file)\b|\bshrink\b', msg):
            return "compress"
        if re.search(r'\brepair\b|\bfix\b.*\b(pdf|file)\b|\bdamage\b|\bcorrupt\b', msg):
            return "repair"

        # SECURITY
        if re.search(r'\bsign\b.*\b(pdf|document|file)\b|\bsignature\b', msg):
            return "sign"
        if re.search(r'\bunlock\b|\bremove\s+password\b|\bdecrypt\b', msg):
            return "unlock"
        if re.search(r'\bprotect\b|\bpassword\b|\bencrypt\b', msg):
            return "protect"
        if re.search(r'\brename\b', msg):
            return "rename"

        # AI
        if re.search(r'\bsummar(ize|y|ise)\b|\btl;?dr\b|\bbrief\b|\boverview\b|\bkey\s+points?\b|\bexecutive\b|\bhighlight', msg):
            return "summarize"
        if re.search(r'\btranslat', msg):
            return "translate"

        return "chat"

    # ─────────────────────────────────────────────
    # CONTEXT RESPONSES (tool briefing from navbar)
    # ─────────────────────────────────────────────

    def _context_response(self, tool: str) -> dict:
        responses = {
            "context_merge": "**Merge PDF activated.** Upload two or more PDF files in the sidebar, then tell me something like:\n- *Merge all my PDFs*\n- *Combine in order: file1, file2, file3*",
            "context_split": "**Split PDF activated.** Upload your PDF and say:\n- *Split the PDF* — each page becomes its own file, bundled in a ZIP.",
            "context_remove_pages": "**Remove Pages activated.** Upload your PDF, then tell me:\n- *Remove pages 3, 5, 7*\n- *Delete pages 10-15*",
            "context_extract": "**Extract Pages activated.** Upload your PDF, then specify:\n- *Extract pages 1-5*\n- *Pull out page 3*",
            "context_organize": "**Organize Pages activated.** Upload your PDF and specify the new page order:\n- *Reorder pages: 3, 1, 2, 4*\n- *Put page 5 first, then 1, 2, 3, 4*",
            "context_rotate": "**Rotate PDF activated.** Upload your PDF and tell me:\n- *Rotate all pages 90 degrees*\n- *Rotate pages 2-4 by 180 degrees*",
            "context_watermark": "**Watermark activated.** Upload your PDF and say:\n- *Add 'CONFIDENTIAL' watermark*\n- *Watermark with the text 'DRAFT'*",
            "context_page_numbers": "**Page Numbers activated.** Upload your PDF and say:\n- *Add page numbers at the bottom center*\n- *Number pages starting from 5, top right*",
            "context_jpg_to_pdf": "**JPG to PDF activated.** Upload your image file(s) (JPG, PNG) in the sidebar, then say:\n- *Convert to PDF*\n- *Make a PDF from my images*",
            "context_word_to_pdf": "**Word to PDF activated.** Upload your .docx file in the sidebar, then say:\n- *Convert to PDF*\n- *Turn my Word document into a PDF*",
            "context_pptx_to_pdf": "**PowerPoint to PDF activated.** Upload your .pptx file in the sidebar, then say:\n- *Convert to PDF*\n- *Make a PDF from my presentation*",
            "context_excel_to_pdf": "**Excel to PDF activated.** Upload your .xlsx file in the sidebar, then say:\n- *Convert to PDF*\n- *Turn my spreadsheet into a PDF*",
            "context_html_to_pdf": "**HTML to PDF activated.** Say:\n- *Convert this HTML: <p>Hello world</p>*\n- I'll render your HTML and produce a PDF.",
            "context_pdf_to_jpg": "**PDF to JPG activated.** Upload your PDF, then say:\n- *Convert to images*\n- *Export pages as JPEG* — I'll create a ZIP with one JPG per page.",
            "context_pdf_to_word": "**PDF to Word activated.** Upload your PDF, then say:\n- *Convert to Word*\n- *Export as DOCX*",
            "context_pdf_to_pptx": "**PDF to PowerPoint activated.** Upload your PDF, then say:\n- *Convert to PowerPoint*\n- *Make a PPTX from this PDF* — each page becomes a slide.",
            "context_pdf_to_excel": "**PDF to Excel activated.** Upload your PDF (works best with tables), then say:\n- *Convert to Excel*\n- *Extract tables to spreadsheet*",
            "context_pdf_to_pdfa": "**PDF to PDF/A activated.** Upload your PDF, then say:\n- *Convert to PDF/A*\n- *Archive this document* — I'll create a long-term archival version.",
            "context_compress": "**Compress PDF activated.** Upload your PDF, then click **Compress PDF** or **Compress as small as possible** below.",
            "context_repair": "**Repair PDF activated.** Upload your damaged or corrupt PDF, then click **Repair this PDF** below.",
            "context_sign": "**Sign PDF activated.** Upload your PDF *and* a signature image (PNG/JPG), then pick a signing position from the quick-start buttons below.",
            "context_unlock": "**Unlock PDF activated.** Upload your password-protected PDF, then click **Unlock PDF** below. If it has a password, include it in your message.",
            "context_protect": "**Protect PDF activated.** Upload your PDF, then click **🔒 Protect with password:** or **🔐 Encrypt with password:** below and type your password at the end.",
            "context_summarize": "**AI Summarizer activated.** Upload your PDF, then click one of the summary options below — Summarize, TL;DR, Key points, or Executive summary.",
            "context_translate": "**AI Translate activated.** Upload your PDF, then click a language button below. Need a different language? Use **Other language…**",
        }
        content = responses.get(tool, "Tool activated. Upload your files and tell me what to do.")
        return {
            "id": str(uuid.uuid4()),
            "role": "assistant",
            "content": content,
            "toolsUsed": [],
            "resultFiles": [],
        }

    # ─────────────────────────────────────────────
    # TOOL EXECUTION
    # ─────────────────────────────────────────────

    async def process_message(self, message: str, file_ids: list[str] = None) -> dict:
        if file_ids is None:
            file_ids = []

        intent = self._detect_intent(message)

        # Context responses (no file processing needed)
        if intent.startswith("context_"):
            return self._context_response(intent)

        try:
            if intent == "merge":
                return await self._handle_merge(file_ids)
            elif intent == "split":
                return await self._handle_split(message, file_ids)
            elif intent == "remove_pages":
                return await self._handle_remove_pages(message, file_ids)
            elif intent == "extract":
                return await self._handle_extract(message, file_ids)
            elif intent == "organize":
                return await self._handle_organize(message, file_ids)
            elif intent == "rotate":
                return await self._handle_rotate(message, file_ids)
            elif intent == "watermark":
                return await self._handle_watermark(message, file_ids)
            elif intent == "page_numbers":
                return await self._handle_page_numbers(message, file_ids)
            elif intent == "jpg_to_pdf":
                return await self._handle_images_to_pdf(file_ids)
            elif intent == "word_to_pdf":
                return await self._handle_word_to_pdf(file_ids)
            elif intent == "pptx_to_pdf":
                return await self._handle_pptx_to_pdf(file_ids)
            elif intent == "excel_to_pdf":
                return await self._handle_excel_to_pdf(file_ids)
            elif intent == "html_to_pdf":
                return await self._handle_html_to_pdf(message, file_ids)
            elif intent == "pdf_to_jpg":
                return await self._handle_pdf_to_jpg(file_ids)
            elif intent == "pdf_to_word":
                return await self._handle_pdf_to_word(file_ids)
            elif intent == "pdf_to_pptx":
                return await self._handle_pdf_to_pptx(file_ids)
            elif intent == "pdf_to_excel":
                return await self._handle_pdf_to_excel(file_ids)
            elif intent == "pdf_to_pdfa":
                return await self._handle_pdf_to_pdfa(file_ids)
            elif intent == "pdf_to_text":
                return await self._handle_pdf_to_text(file_ids)
            elif intent == "compress":
                return await self._handle_compress(file_ids)
            elif intent == "repair":
                return await self._handle_repair(file_ids)
            elif intent == "sign":
                return await self._handle_sign(message, file_ids)
            elif intent == "unlock":
                return await self._handle_unlock(message, file_ids)
            elif intent == "protect":
                return await self._handle_protect(message, file_ids)
            elif intent == "rename":
                return await self._handle_rename(message, file_ids)
            elif intent == "summarize":
                return await self._handle_summarize(file_ids, message)
            elif intent == "translate":
                return await self._handle_translate(message, file_ids)
            else:
                return await self._handle_chat(message, file_ids)
        except Exception as e:
            return {
                "id": str(uuid.uuid4()),
                "role": "assistant",
                "content": f"⚠️ Error while processing: {str(e)}\n\nPlease make sure the correct file is uploaded and try again.",
                "toolsUsed": [],
                "resultFiles": [],
            }

    # ─────────────────────────────────────────────
    # HANDLERS
    # ─────────────────────────────────────────────

    async def _handle_merge(self, file_ids):
        pdfs = self._all_pdfs(file_ids)
        if len(pdfs) < 2:
            return self._msg("Please upload **at least 2 PDF files** to merge.", [])
        paths = [f["path"] for f in pdfs]
        out_path, out_name = self.tools.merge_pdfs(paths)
        result = self._store_result(out_path, out_name)
        names = ", ".join(f["name"] for f in pdfs)
        return self._msg(
            f"✅ Merged **{len(pdfs)} PDFs** ({names}) into one file.",
            ["merge_pdfs"], [result]
        )

    async def _handle_split(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to split.", [])
        # Check for "split at page N" / "split at N"
        at_match = re.search(r'(?:at|after|from)\s+page\s+(\d+)|split\s+(?:at\s+)?page\s+(\d+)', message, re.IGNORECASE)
        if at_match:
            split_page = int(at_match.group(1) or at_match.group(2))
            out_path, out_name = self.tools.split_at_page(pdf["path"], split_page)
            result = self._store_result(out_path, out_name)
            return self._msg(
                f"✅ Split **{pdf['name']}** into 2 parts at page {split_page}. Both parts are in the ZIP.",
                ["split_pdf"], [result]
            )
        out_path, out_name = self.tools.split_pdf(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Split **{pdf['name']}** — each page is now a separate PDF inside the downloaded ZIP.",
            ["split_pdf"], [result]
        )

    async def _handle_remove_pages(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to remove pages from.", [])

        import fitz as _fitz
        with _fitz.open(pdf["path"]) as _doc:
            total = len(_doc)

        msg = message.lower()
        # Semantic shortcuts
        if re.search(r'first\s+page|page\s+1\b', msg) and not re.search(r'\d', msg.replace("page 1", "")):
            pages = [1]
        elif re.search(r'last\s+page', msg):
            pages = [total]
        else:
            pages = self._parse_page_list(message)

        if not pages:
            return self._msg(
                "Please specify which pages to remove. Examples:\n"
                "- *Remove the first page*\n"
                "- *Remove the last page*\n"
                "- *Remove pages 3, 5, 7*",
                []
            )
        out_path, out_name = self.tools.remove_pages(pdf["path"], pages)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Removed page{'s' if len(pages) > 1 else ''} **{', '.join(str(p) for p in pages)}** from **{pdf['name']}**.",
            ["remove_pages"], [result]
        )

    async def _handle_extract(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to extract pages from.", [])

        import fitz as _fitz
        with _fitz.open(pdf["path"]) as _doc:
            total = len(_doc)

        msg = message.lower()
        # Semantic shortcuts
        if re.search(r'first\s+(?:3|three)\s+pages?', msg):
            page_range = (1, min(3, total))
        elif re.search(r'first\s+page|page\s+1\b', msg) and not re.search(r'\d', msg.replace("page 1", "")):
            page_range = (1, 1)
        elif re.search(r'last\s+page', msg):
            page_range = (total, total)
        elif re.search(r'last\s+(?:3|three)\s+pages?', msg):
            page_range = (max(1, total - 2), total)
        else:
            page_range = self._parse_page_range(message)

        if not page_range:
            return self._msg(
                "Please specify pages to extract. Examples:\n"
                "- *Extract page 1*\n"
                "- *Extract the last page*\n"
                "- *Extract pages 2–5*",
                []
            )
        out_path, out_name = self.tools.extract_pages(pdf["path"], page_range[0], page_range[1])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Extracted pages **{page_range[0]}–{page_range[1]}** from **{pdf['name']}**.",
            ["extract_pages"], [result]
        )

    async def _handle_organize(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to organize.", [])

        # Get total page count for semantic commands
        import fitz as _fitz
        with _fitz.open(pdf["path"]) as _doc:
            total = len(_doc)

        msg = message.lower()
        # Semantic commands — no page numbers needed
        if re.search(r'revers', msg):
            order = list(range(total, 0, -1))
        elif re.search(r'first.*(to|at).*(end|last|back)', msg) or re.search(r'move.*(first|page 1).*(end|last)', msg):
            order = list(range(2, total + 1)) + [1]
        elif re.search(r'last.*(to|at).*(front|first|top)', msg) or re.search(r'move.*last.*(front|first)', msg):
            order = [total] + list(range(1, total))
        else:
            order = self._parse_page_list(message, preserve_order=True)

        if not order:
            return self._msg(
                "Please specify the page order. Examples:\n"
                "- *Reverse the page order*\n"
                "- *Move the last page to the front*\n"
                "- *Reorder pages: 3, 1, 2, 4*",
                []
            )
        out_path, out_name = self.tools.organize_pages(pdf["path"], order)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Reordered pages in **{pdf['name']}** to: {', '.join(str(p) for p in order)}.",
            ["organize_pages"], [result]
        )

    async def _handle_rotate(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to rotate.", [])
        angle_match = re.search(r'(\d+)\s*(?:degree|°)', message, re.IGNORECASE)
        angle = int(angle_match.group(1)) if angle_match else 90
        angle = round(angle / 90) * 90  # Snap to 90° increments
        # Strip the angle value from the message before parsing page numbers
        msg_for_pages = re.sub(r'\d+\s*(?:degree|°)', '', message, flags=re.IGNORECASE)
        msg_lower = message.lower()

        # Semantic page shortcuts
        import fitz as _fitz
        with _fitz.open(pdf["path"]) as _doc:
            total_pages = len(_doc)

        if re.search(r'first\s+page', msg_lower):
            pages = [1]
        elif re.search(r'last\s+page', msg_lower):
            pages = [total_pages]
        else:
            # Only parse explicit "page N" / "pages N-M" patterns (not bare numbers)
            page_match = re.search(r'pages?\s+([\d,\s\-–]+)', msg_for_pages, re.IGNORECASE)
            if page_match:
                pages = self._parse_page_list(page_match.group(1)) or None
            else:
                pages = None
        out_path, out_name = self.tools.rotate_pages(pdf["path"], angle, pages)
        result = self._store_result(out_path, out_name)
        target = f"pages {', '.join(str(p) for p in pages)}" if pages else "all pages"
        return self._msg(
            f"✅ Rotated **{target}** of **{pdf['name']}** by **{angle}°**.",
            ["rotate_pages"], [result]
        )

    async def _handle_watermark(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to watermark.", [])
        # Match explicit "text: X", "watermark: X", or "stamp: X" patterns
        text_match = re.search(r'(?:text|label|watermark|stamp)\s*:\s*[\'"]?([A-Za-z0-9 _\-]+)[\'"]?', message, re.IGNORECASE)
        if not text_match:
            # Fallback: word following "watermark with" or "stamp with"
            text_match = re.search(r'(?:watermark|stamp)\s+(?:with\s+)?[\'"]([A-Za-z0-9 _\-]+)[\'"]', message, re.IGNORECASE)
        wm_text = text_match.group(1).strip().upper() if text_match else "CONFIDENTIAL"
        out_path, out_name = self.tools.add_watermark(pdf["path"], text=wm_text)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Added **'{wm_text}'** watermark to all pages of **{pdf['name']}**.",
            ["add_watermark"], [result]
        )

    async def _handle_page_numbers(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to add page numbers to.", [])
        msg_lower = message.lower()
        if "top right" in msg_lower:
            position = "top-right"
        elif "top left" in msg_lower:
            position = "top-left"
        elif "top" in msg_lower:
            position = "top-center"
        elif "bottom right" in msg_lower:
            position = "bottom-right"
        elif "bottom left" in msg_lower:
            position = "bottom-left"
        else:
            position = "bottom-center"
        start_match = re.search(r'start(?:ing)?\s+(?:from|at)?\s*(\d+)', message, re.IGNORECASE)
        start_num = int(start_match.group(1)) if start_match else 1
        out_path, out_name = self.tools.add_page_numbers(pdf["path"], position=position, start_num=start_num)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Added page numbers at **{position}** to **{pdf['name']}** (starting from {start_num}).",
            ["add_page_numbers"], [result]
        )

    async def _handle_images_to_pdf(self, file_ids):
        files = self._get_input_files(file_ids)
        images = [f for f in files if f["path"].lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"))]
        if not images:
            return self._msg("Please upload **image files** (JPG, PNG, etc.) to convert to PDF.", [])
        paths = [f["path"] for f in images]
        out_path, out_name = self.tools.images_to_pdf(paths)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{len(images)} image(s)** to a single PDF.",
            ["images_to_pdf"], [result]
        )

    async def _handle_word_to_pdf(self, file_ids):
        files = self._get_input_files(file_ids)
        docx_files = [f for f in files if f["path"].lower().endswith((".docx", ".doc"))]
        if not docx_files:
            return self._msg("Please upload a **Word document (.docx)** to convert to PDF.", [])
        out_path, out_name = self.tools.word_to_pdf(docx_files[0]["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{docx_files[0]['name']}** to PDF.",
            ["word_to_pdf"], [result]
        )

    async def _handle_pptx_to_pdf(self, file_ids):
        files = self._get_input_files(file_ids)
        pptx_files = [f for f in files if f["path"].lower().endswith((".pptx", ".ppt"))]
        if not pptx_files:
            return self._msg("Please upload a **PowerPoint file (.pptx)** to convert to PDF.", [])
        out_path, out_name = self.tools.pptx_to_pdf(pptx_files[0]["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{pptx_files[0]['name']}** to PDF.",
            ["pptx_to_pdf"], [result]
        )

    async def _handle_excel_to_pdf(self, file_ids):
        files = self._get_input_files(file_ids)
        xlsx_files = [f for f in files if f["path"].lower().endswith((".xlsx", ".xls"))]
        if not xlsx_files:
            return self._msg("Please upload an **Excel file (.xlsx)** to convert to PDF.", [])
        out_path, out_name = self.tools.excel_to_pdf(xlsx_files[0]["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{xlsx_files[0]['name']}** to PDF.",
            ["excel_to_pdf"], [result]
        )

    async def _handle_html_to_pdf(self, message, file_ids):
        html_match = re.search(r'(<[a-z][\s\S]*>[\s\S]*<\/[a-z]+>|<[a-z][^>]*\/>)', message, re.IGNORECASE)
        if not html_match:
            return self._msg(
                "Please include HTML content in your message. Example:\n"
                "*Convert this HTML: `<h1>Hello</h1><p>World</p>`*",
                []
            )
        html_content = message[html_match.start():]
        out_path, out_name = self.tools.html_to_pdf(html_content)
        result = self._store_result(out_path, out_name)
        return self._msg("✅ Converted your HTML to PDF.", ["html_to_pdf"], [result])

    async def _handle_pdf_to_jpg(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to convert to images.", [])
        out_path, out_name = self.tools.pdf_to_jpg(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{pdf['name']}** to JPEG images. Download the ZIP — one image per page.",
            ["pdf_to_jpg"], [result]
        )

    async def _handle_pdf_to_word(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to convert to Word.", [])
        out_path, out_name = self.tools.pdf_to_word(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{pdf['name']}** to a Word document (.docx) with preserved text structure.",
            ["pdf_to_word"], [result]
        )

    async def _handle_pdf_to_pptx(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to convert to PowerPoint.", [])
        out_path, out_name = self.tools.pdf_to_pptx(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{pdf['name']}** to PowerPoint — each PDF page is a slide.",
            ["pdf_to_pptx"], [result]
        )

    async def _handle_pdf_to_excel(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to convert to Excel.", [])
        out_path, out_name = self.tools.pdf_to_excel(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Extracted tables from **{pdf['name']}** into an Excel spreadsheet.",
            ["pdf_to_excel"], [result]
        )

    async def _handle_pdf_to_pdfa(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to convert to PDF/A.", [])
        out_path, out_name = self.tools.pdf_to_pdfa(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Converted **{pdf['name']}** to archival PDF/A format.",
            ["pdf_to_pdfa"], [result]
        )

    async def _handle_pdf_to_text(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to extract text from.", [])
        out_path, out_name = self.tools.pdf_to_text(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Extracted all text from **{pdf['name']}** to a plain text file.",
            ["pdf_to_text"], [result]
        )

    async def _handle_compress(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to compress.", [])
        import os
        original_size = os.path.getsize(pdf["path"])
        out_path, out_name = self.tools.compress_pdf(pdf["path"])
        result = self._store_result(out_path, out_name)
        new_size = os.path.getsize(out_path)
        reduction = round((1 - new_size / original_size) * 100, 1) if original_size > 0 else 0
        size_str = f"{new_size / 1024:.0f} KB" if new_size < 1024 * 1024 else f"{new_size / (1024*1024):.2f} MB"
        return self._msg(
            f"✅ Compressed **{pdf['name']}** — {reduction}% smaller ({size_str}).",
            ["compress_pdf"], [result]
        )

    async def _handle_repair(self, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload the **damaged PDF file** you want to repair.", [])
        out_path, out_name = self.tools.repair_pdf(pdf["path"])
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Repaired **{pdf['name']}** — the recovered and cleaned PDF is ready to download.",
            ["repair_pdf"], [result]
        )

    async def _handle_sign(self, message, file_ids):
        files = self._get_input_files(file_ids)
        pdfs = [f for f in files if f["path"].lower().endswith(".pdf")]
        sigs = [f for f in files if f["path"].lower().endswith((".png", ".jpg", ".jpeg"))]
        if not pdfs:
            return self._msg("Please upload a **PDF file** to sign.", [])
        if not sigs:
            return self._msg("Please also upload a **signature image** (PNG or JPG) in the sidebar.", [])
        msg_lower = message.lower()
        position = "top" if "top" in msg_lower else ("center" if "center" in msg_lower else "bottom")
        page_target = "all" if "all page" in msg_lower else ("first" if "first" in msg_lower else "last")
        out_path, out_name = self.tools.sign_pdf(pdfs[0]["path"], sigs[0]["path"], position, page_target)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Signed **{pdfs[0]['name']}** with your signature at the **{position}** of the **{page_target}** page(s).",
            ["sign_pdf"], [result]
        )

    async def _handle_unlock(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload the **password-protected PDF** you want to unlock.", [])
        password_match = re.search(r'password[:\s]+["\']?(\S+)["\']?', message, re.IGNORECASE)
        password = password_match.group(1) if password_match else ""
        out_path, out_name = self.tools.unlock_pdf(pdf["path"], password)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Unlocked **{pdf['name']}** — the password protection has been removed.",
            ["unlock_pdf"], [result]
        )

    async def _handle_protect(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to protect.", [])
        password_match = re.search(r'password[:\s]+["\']?(\S+)["\']?', message, re.IGNORECASE)
        if not password_match:
            return self._msg("Please specify a password. Example: *Protect with password: mysecret123*", [])
        password = password_match.group(1)
        out_path, out_name = self.tools.protect_pdf(pdf["path"], password)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Protected **{pdf['name']}** with a password. Keep it safe!",
            ["protect_pdf"], [result]
        )

    async def _handle_rename(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to rename.", [])
        name_match = re.search(r"(?:rename|call|name)\s+(?:it\s+)?[\"']?([^\"']+)[\"']?$", message, re.IGNORECASE)
        if not name_match:
            return self._msg("Please specify the new name. Example: *Rename to Final_Report*", [])
        new_name = name_match.group(1).strip()
        out_path, out_name = self.tools.rename_pdf(pdf["path"], new_name)
        result = self._store_result(out_path, out_name)
        return self._msg(
            f"✅ Renamed **{pdf['name']}** to **{out_name}**.",
            ["rename_pdf"], [result]
        )

    async def _handle_summarize(self, file_ids, message=""):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to summarize.", [])
        text = self.tools.extract_full_text(pdf["path"])
        if not text.strip():
            return self._msg("The PDF appears to be empty or image-only (no extractable text).", [])

        # Tailor prompt based on request type
        msg_lower = message.lower()
        if re.search(r'tl;?dr', msg_lower):
            system_prompt = "You are a helpful assistant. Respond with a single short paragraph (2-3 sentences max) as a TL;DR."
            user_prompt = f"Give me a TL;DR of this document:\n\n{text}"
            label = "⚡ TL;DR"
        elif re.search(r'key\s+point', msg_lower):
            system_prompt = "You are a helpful assistant. Extract only the most important key points as a bullet list. Be concise."
            user_prompt = f"Extract the key points from this document:\n\n{text}"
            label = "🔑 Key Points"
        elif re.search(r'executive', msg_lower):
            system_prompt = "You are a business analyst. Write a professional executive summary with: Overview, Key Findings, and Recommendations sections."
            user_prompt = f"Write an executive summary of this document:\n\n{text}"
            label = "📊 Executive Summary"
        else:
            system_prompt = "You are a helpful assistant that summarizes documents. Be concise and use bullet points for key insights."
            user_prompt = f"Summarize this document:\n\n{text}"
            label = "📄 Summary"

        if self._openai_available:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                max_tokens=700,
            )
            summary = response.choices[0].message.content
        else:
            # Fallback: extract key sentences
            sentences = [s.strip() for s in text.replace("\n", " ").split(".") if len(s.strip()) > 40]
            summary = "**Key sentences:**\n\n" + "\n".join(f"• {s}." for s in sentences[:8])
        return {
            "id": str(uuid.uuid4()),
            "role": "assistant",
            "content": f"**{label} — {pdf['name']}:**\n\n{summary}",
            "toolsUsed": ["ai_summarize"],
            "resultFiles": [],
        }

    async def _handle_translate(self, message, file_ids):
        pdf = self._first_pdf(file_ids)
        if not pdf:
            return self._msg("Please upload a **PDF file** to translate.", [])
        lang_match = re.search(r'to\s+([A-Za-z]+(?:\s+[A-Za-z]+)?)', message, re.IGNORECASE)
        target_lang = lang_match.group(1).strip().title() if lang_match else "Spanish"
        text = self.tools.extract_full_text(pdf["path"])
        if not text.strip():
            return self._msg("The PDF appears to be empty or image-only (no extractable text to translate).", [])
        if not self._openai_available:
            return self._msg(
                f"AI translation to **{target_lang}** requires an OpenAI API key. "
                "Please set the OPENAI_API_KEY environment variable.",
                []
            )
        # Translate in chunks if needed (already capped at 12K chars)
        response = self.client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": f"You are a professional translator. Translate the given text to {target_lang}. Preserve formatting and structure."},
                {"role": "user", "content": text[:6000]},
            ],
            max_tokens=2000,
        )
        translated = response.choices[0].message.content
        # Save translation as a text file
        output_name = f"{Path(pdf['path']).stem}_translated_{target_lang.lower()}_{uuid.uuid4().hex[:6]}.txt"
        output_path = str(self.temp_dir / output_name)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"Translation to {target_lang}\nOriginal: {pdf['name']}\n\n")
            f.write(translated)
        result = self._store_result(output_path, output_name)
        return {
            "id": str(uuid.uuid4()),
            "role": "assistant",
            "content": f"🌐 Translated **{pdf['name']}** to **{target_lang}**. Preview:\n\n{translated[:500]}{'...' if len(translated) > 500 else ''}",
            "toolsUsed": ["ai_translate"],
            "resultFiles": [result],
        }

    async def _handle_chat(self, message, file_ids):
        files = self._get_input_files(file_ids)
        file_context = ""
        if files:
            file_context = "Uploaded files: " + ", ".join(f["name"] for f in files) + "\n\n"
        if self._openai_available:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": (
                        "You are a helpful PDF AI assistant. You can help with merging, splitting, "
                        "extracting, rotating, watermarking, compressing, converting, and signing PDF files. "
                        "Guide the user clearly on what to do with their files."
                    )},
                    {"role": "user", "content": file_context + message},
                ],
                max_tokens=400,
            )
            content = response.choices[0].message.content
        else:
            content = (
                "I'm your PDF AI assistant. I can help you with:\n\n"
                "**Organize:** Merge, Split, Extract Pages, Remove Pages, Reorder Pages, Rotate\n"
                "**Convert to PDF:** JPG, Word, PowerPoint, Excel, HTML\n"
                "**Convert from PDF:** to JPG, Word, PowerPoint, Excel, PDF/A, Text\n"
                "**Optimize:** Compress, Repair, Watermark, Page Numbers\n"
                "**Security:** Sign, Unlock, Protect with Password\n"
                "**AI:** Summarize, Translate\n\n"
                "Upload your files in the sidebar and tell me what to do!"
            )
        return {
            "id": str(uuid.uuid4()),
            "role": "assistant",
            "content": content,
            "toolsUsed": [],
            "resultFiles": [],
        }

    def _msg(self, content: str, tools_used: list, result_files: list = None) -> dict:
        return {
            "id": str(uuid.uuid4()),
            "role": "assistant",
            "content": content,
            "toolsUsed": tools_used,
            "resultFiles": result_files or [],
        }
