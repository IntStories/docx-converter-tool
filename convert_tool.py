import os
import shutil
import zipfile
import tempfile
import pypandoc
from docx import Document
from docx2pdf import convert as docx2pdf_convert
import sys
from tkinter import Tk, filedialog

# Set the path to the bundled pandoc.exe
def set_local_pandoc():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    pandoc_path = os.path.join(base_path, 'pandoc', 'pandoc.exe')
    if not os.path.isfile(pandoc_path):
        print("❌ pandoc.exe not found. Please ensure it's included in the 'pandoc' folder.")
        sys.exit(1)
    pypandoc.pandoc_path = pandoc_path

# Convert DOCX to a spaced-out TXT
def convert_to_txt_with_spacing(docx_path, output_path):
    doc = Document(docx_path)
    with open(output_path, 'w', encoding='utf-8') as f:
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                f.write(text + "\n\n")

# Open file dialog to choose the DOCX file
def pick_docx_file():
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    return filedialog.askopenfilename(
        title="Select a .docx file",
        filetypes=[("Word Documents", "*.docx")]
    )

# Save as dialog to choose where to save the ZIP
def pick_save_location(default_name):
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    return filedialog.asksaveasfilename(
        title="Save ZIP File As",
        defaultextension=".zip",
        filetypes=[("ZIP Archives", "*.zip")],
        initialfile=default_name + "_bundle.zip"
    )

def main():
    print("📄 DOCX Conversion & Zipping Tool\n")
    set_local_pandoc()

    docx_path = pick_docx_file()
    if not docx_path or not os.path.isfile(docx_path) or not docx_path.lower().endswith('.docx'):
        print("❌ Invalid or no .docx file selected.")
        return

    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    temp_dir = tempfile.mkdtemp()
    print(f"🔧 Working in temporary folder: {temp_dir}")

    shutil.copy(docx_path, os.path.join(temp_dir, f"{base_name}.docx"))

    targets = {
        "html": f"{base_name}.html",
        "odt": f"{base_name}.odt",
        "epub": f"{base_name}.epub",
        "pdf": f"{base_name}.pdf"
    }

    for fmt, filename in targets.items():
        output_path = os.path.join(temp_dir, filename)
        try:
            if fmt == "pdf":
                docx2pdf_convert(docx_path, temp_dir)
                print("✅ Converted to PDF (preserving original style)")
            else:
                pypandoc.convert_file(docx_path, fmt, outputfile=output_path)
                print(f"✅ Converted to {fmt.upper()}")
        except Exception as e:
            print(f"❌ Failed to convert to {fmt.upper()}: {e}")

    txt_path = os.path.join(temp_dir, f"{base_name}.txt")
    try:
        convert_to_txt_with_spacing(docx_path, txt_path)
        print("✅ Converted to TXT with paragraph spacing")
    except Exception as e:
        print(f"❌ Failed to convert to TXT: {e}")

    zip_path = pick_save_location(base_name)
    if not zip_path:
        print("❌ No save location selected.")
        shutil.rmtree(temp_dir)
        return

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(temp_dir):
            full_path = os.path.join(temp_dir, file)
            zipf.write(full_path, arcname=file)

    print(f"\n📦 Zip created: {zip_path}")

    shutil.rmtree(temp_dir)
    print("🧹 Cleaned up temporary files.")

if __name__ == "__main__":
    main()
