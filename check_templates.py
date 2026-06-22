from pathlib import Path
import zipfile
import re

def extract_merge_fields(docx_path):
    """Extract all {{field}} from a .docx file"""
    try:
        fields = set()
        with zipfile.ZipFile(docx_path) as docx:
            # Check main document
            if 'word/document.xml' in docx.namelist():
                xml = docx.read('word/document.xml').decode('utf-8', errors='ignore')
                matches = re.findall(r'\{\{([^{}]+?)\}\}', xml)
                fields.update(matches)

            # Check headers and footers
            for name in docx.namelist():
                if name.startswith('word/header') or name.startswith('word/footer'):
                    xml = docx.read(name).decode('utf-8', errors='ignore')
                    matches = re.findall(r'\{\{([^{}]+?)\}\}', xml)
                    fields.update(matches)

        return sorted(list(fields))
    except Exception as e:
        return [f"ERROR reading file: {e}"]

# Main
print("🔍 Scanning all templates for merge fields...\n")

templates_dir = Path("templates")
all_fields = set()

for docx_file in sorted(templates_dir.glob("*.docx")):
    print(f"📄 {docx_file.name}")
    fields = extract_merge_fields(docx_file)
    print(f"   → {len(fields)} fields: {fields}\n")
    all_fields.update(fields)

print("=" * 70)
print("📋 SUMMARY - ALL UNIQUE MERGE FIELDS FOUND:")
print("=" * 70)
for field in sorted(all_fields):
    print(f"• {field}")

print(f"\nTotal unique merge fields: {len(all_fields)}")
print("Done.")
