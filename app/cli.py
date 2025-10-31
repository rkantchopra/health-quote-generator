import argparse
from app.processor import generate_docx

def main():
    p = argparse.ArgumentParser(description="Generate Health Quote DOCX from Excel")
    p.add_argument("excel", help="Path to input Excel with sheets 'Client Details' and 'Premiums'")
    p.add_argument("-o", "--output", default="output/Health_Quote.docx", help="Output DOCX path")
    p.add_argument("-l", "--logos", default="logos", help="Folder containing insurer logos")
    args = p.parse_args()

    out = generate_docx(args.excel, args.output, args.logos)
    print(f"✅ Generated: {out}")

if __name__ == "__main__":
    main()
