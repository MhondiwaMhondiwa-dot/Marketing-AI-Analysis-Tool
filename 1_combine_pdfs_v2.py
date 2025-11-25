#!/usr/bin/env python3
"""
Combines all PDFs in the folder into one file with a Table of Contents.
v2: Adds progress printing so you know it's working.
"""
from __future__ import annotations
import io
import re
from pathlib import Path
from typing import List

from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def build_visual_toc(entries: List[dict]) -> PdfReader:
    """Generates the visual page for the Table of Contents."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setTitle("Table of Contents")
    
    # Title
    c.setFont("Helvetica-Bold", 24)
    c.drawString(72, 750, "Table of Contents")
    
    # Entries
    c.setFont("Helvetica", 12)
    y = 700
    
    for chapter in entries:
        if y < 72:
            c.showPage()
            c.setFont("Helvetica", 12)
            y = 750
            
        label = chapter['name']
        # Truncate long names so they don't overlap page numbers
        if len(label) > 60:
            label = label[:57] + "..."
            
        page_num = str(chapter['start_page'])
        
        c.drawString(72, y, label)
        
        # Draw dotted line
        dot_width = c.stringWidth(".", "Helvetica", 12)
        text_end = 72 + c.stringWidth(label, "Helvetica", 12) + 5
        num_start = 540 - c.stringWidth(page_num, "Helvetica", 12) - 5
        
        current_x = text_end
        while current_x < num_start:
            c.drawString(current_x, y, ".")
            current_x += dot_width + 2

        c.drawRightString(540, y, page_num)
        y -= 20
        
    c.save()
    buffer.seek(0)
    return PdfReader(buffer)

def main():
    pdf_dir = Path.cwd()
    output_filename = "Combined_Assignment_Document.pdf"

    def _numeric_key(path: Path):
        m = re.search(r"(\d+)", path.stem)
        if m:
            return (0, int(m.group(1)), path.stem.lower())
        return (1, float("inf"), path.stem.lower())

    pdf_files = sorted(
        (p for p in pdf_dir.glob("*.pdf") if p.is_file() and p.name != output_filename), 
        key=_numeric_key
    )

    if not pdf_files:
        print("❌ No PDF files found.")
        return

    print(f"📂 Found {len(pdf_files)} PDFs.")

    writer = PdfWriter()
    chapter_entries = []
    current_page_count = 0
    total_files = len(pdf_files)

    # 1. Merge Loop with Progress Bar
    print("-" * 40)
    for idx, pdf_path in enumerate(pdf_files, start=1):
        # Print progress!
        print(f"[{idx}/{total_files}] Processing: {pdf_path.name}...")
        
        try:
            reader = PdfReader(str(pdf_path))
            
            # Skip encrypted/unreadable files
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except:
                    print(f"   ⚠️ Skipped (Encrypted): {pdf_path.name}")
                    continue

            start_page = current_page_count + 1 
            
            chapter_entries.append({
                "name": pdf_path.stem.replace("_", " ").title(),
                "start_page": start_page
            })

            for page in reader.pages:
                writer.add_page(page)
                current_page_count += 1
                
        except Exception as e:
            print(f"   ❌ Error reading {pdf_path.name}: {e}")

    print("-" * 40)
    print("📚 Generating Table of Contents...")

    # 2. Generate TOC
    if not chapter_entries:
        print("❌ No valid pages were added. Exiting.")
        return

    temp_toc = build_visual_toc(chapter_entries)
    toc_num_pages = len(temp_toc.pages)

    for entry in chapter_entries:
        entry['start_page'] += toc_num_pages

    final_toc_reader = build_visual_toc(chapter_entries)

    for page in reversed(final_toc_reader.pages):
        writer.insert_page(page, index=0)

    # 3. Add Bookmarks
    for entry in chapter_entries:
        writer.add_outline_item(entry['name'], entry['start_page'] - 1)

    print("💾 Saving final file... (Do not close!)")
    
    with open(output_filename, "wb") as f:
        writer.write(f)

    print(f"✅ DONE! Created: {output_filename}")

if __name__ == "__main__":
    main()