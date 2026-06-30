"""Regression test: 5 e-ticket PDFs + 2 ride-hailing PDFs + existing ride-hailing baseline."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import importlib.util
spec = importlib.util.spec_from_file_location("app", os.path.join(os.path.dirname(__file__), "app.py"))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

# Test files
test_cases = [
    # ===== 5 e-ticket PDFs (user's main complaint) =====
    ("1521692735_00_机票预订出行订单尾号09253345-电子行程单-李清博.pdf", "机票行程单"),
    ("1521692852_00_订单1128148785017825-电子行程单-李清博.pdf", "机票行程单"),
    ("1521692854_00_订单1128148785017825-电子行程单-李清博.pdf", "机票行程单"),
    ("1521692858_00_英文版机票行程单.pdf", "机票行程单-OTA"),
    ("1521692858_01_中英版机票行程单.pdf", "机票行程单-OTA"),
    # ===== 2 ride-hailing PDFs (regression case - just fixed) =====
    ("1521692850_01_如祺出行电子发票.pdf", "网约车发票"),
    ("1521692851_01_哈啰打车电子发票.pdf", "网约车发票"),
    # ===== Existing ride-hailing baseline (must not regress) =====
    ("1521692847_00_滴滴出行行程报销单A.pdf", "网约车发票"),
    ("1521692847_03_滴滴电子发票A.pdf", "网约车发票"),
    ("1521692848_00_阳光出行电子发票.pdf", "网约车发票"),
    ("1521692849_01_上汽享道出行电子发票.pdf", "网约车发票"),
]

raw_dir = "/var/folders/l7/fhmqhfrj3b3013281vfw1pd40000gn/T/invoice_1f66af52ce20_fke87agh/raw"

def get_text(fpath):
    """Try pdfplumber first, fall back to OCR."""
    import pdfplumber
    text = ""
    try:
        with pdfplumber.open(fpath) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                text += t + "\n"
    except Exception as e:
        text = ""
    if len(text.strip()) < 30:
        # OCR fallback
        try:
            text = app._ocr_pdf_with_vision(fpath, dpi=400)
        except Exception as e:
            text = f"OCR_ERROR: {e}"
    return text

print(f"{'FILE':<70} {'EXPECT':<20} {'GOT':<20} {'STATUS'}")
print("=" * 130)
pass_count = 0
fail_count = 0

for filename, expected_subtype in test_cases:
    fpath = os.path.join(raw_dir, filename)
    if not os.path.exists(fpath):
        print(f"{filename:<70} {'(file not found)':<20}")
        fail_count += 1
        continue

    text = get_text(fpath)
    kind, subtype = app.classify_file(filename, text)

    if subtype == expected_subtype:
        status = "✓ PASS"
        pass_count += 1
    else:
        status = f"✗ FAIL (got {kind}/{subtype})"
        fail_count += 1

    print(f"{filename:<70} {expected_subtype:<20} {subtype:<20} {status}")

    # Also try to parse if invoice
    if kind == 'invoice':
        try:
            if subtype == '机票行程单':
                seller, date, amount, remark = app.parse_air_itinerary(text)
                print(f"  → air_itinerary: seller={seller!r} date={date!r} amount={amount!r}")
            elif subtype == '机票行程单-OTA':
                seller, date, amount, remark = app.parse_ota_itinerary(text, filename)
                print(f"  → ota: seller={seller!r} date={date!r} amount={amount!r}")
            elif subtype == '网约车发票':
                seller, date, amount, remark = app.parse_didi_receipt(text)
                print(f"  → ride_hailing: seller={seller!r} date={date!r} amount={amount!r}")
        except Exception as e:
            print(f"  → parse error: {e}")

print()
print(f"Total: {pass_count} PASS, {fail_count} FAIL (of {len(test_cases)})")
sys.exit(0 if fail_count == 0 else 1)
