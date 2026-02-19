
try:
    import pypdf
    print("pypdf is available")
    reader = pypdf.PdfReader("d:/APRENDIZADO APP/MEDICOES/Xerox Scan_19022026162302.PDF")
    print(f"Number of pages: {len(reader.pages)}")
    if len(reader.pages) > 0:
        print("--- Page 1 Content ---")
        print(reader.pages[0].extract_text())
except ImportError:
    print("pypdf not found")
except Exception as e:
    print(f"Error reading PDF: {e}")
