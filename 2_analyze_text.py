#!/usr/bin/env python3
"""
Extracts text from the combined PDF, performs a keyword frequency analysis,
and saves the raw text to a file for easier summarization.
"""
from pypdf import PdfReader
from collections import Counter
import re
from pathlib import Path

def analyze_pdf():
    input_filename = "Combined_Assignment_Document.pdf"
    
    if not Path(input_filename).exists():
        print(f"❌ Could not find {input_filename}. Run the merger script first!")
        return

    print("Reading PDF and extracting text... (This might take a moment)")
    reader = PdfReader(input_filename)
    full_text = []
    
    # Extract text page by page
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text.append(text)

    all_text_content = "\n".join(full_text)
    
    # --- ANALYSIS SECTION ---
    
    # 1. Clean text (remove punctuation, make lowercase)
    words = re.findall(r'\w+', all_text_content.lower())
    
    # 2. Filter out common "stop words" (the, and, is, etc.) so analysis is useful
    stop_words = {
        'the', 'and', 'of', 'to', 'in', 'a', 'is', 'that', 'for', 'it', 'as', 'with', 
        'on', 'are', 'this', 'by', 'be', 'or', 'at', 'from', 'an', 'not', 'can', 'which'
    }
    filtered_words = [w for w in words if w not in stop_words and len(w) > 3]
    
    # 3. Count frequencies
    word_counts = Counter(filtered_words)
    common_words = word_counts.most_common(10)

    # --- OUTPUT SECTION ---

    print("\n" + "="*40)
    print(f"📄 AUTOMATED ANALYSIS REPORT")
    print("="*40)
    print(f"Total Pages Scanned: {len(reader.pages)}")
    print(f"Total Words Counted: {len(words)}")
    print("-" * 30)
    print("TOP 10 KEYWORDS (Subject Matter Hints):")
    for word, count in common_words:
        print(f" • {word.title()}: {count} times")
    print("="*40)

    # Save to file for summarization
    output_txt = "text_for_summary.txt"
    with open(output_txt, "w", encoding="utf-8") as f:
        f.write(all_text_content)
        
    print(f"\n✅ Full text extracted to '{output_txt}'.")
    print("👉 To get your summary: Open this text file, copy all text,")
    print("   and paste it into an AI tool asking for a 'Summary of this text'.")

if __name__ == "__main__":
    analyze_pdf()