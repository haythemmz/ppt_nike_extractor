#!/usr/bin/env python
"""Debug script to test price extraction with specific examples"""

from extractor import extract_prices

# Enable debug mode by modifying extractor temporarily
# For now, just test with our examples

print("=" * 60)
print("TESTING PRICE EXTRACTION")
print("=" * 60)

test_cases = [
    ("Gloves with slash", "Wholesale: $13.50 / Retail: $27.00", (13.50, 27.00)),
    ("Bags newline", "Wholesale: $143.00\nRetail: $260.00", (143.00, 260.00)),
    ("Bags space", "Wholesale: $143.00 Retail: $260.00", (143.00, 260.00)),
    ("Full bag code", "N.101.4771.320\nWholesale: $143.00\nRetail: $260.00", (143.00, 260.00)),
    ("No labels", "$143.00 $260.00", (143.00, 260.00)),
    ("Apparel style", "$70.00 WHOLESALE | $140.00 RETAIL | NEW", (70.00, 140.00)),
]

for name, text, expected in test_cases:
    print(f"\n{name}:")
    print(f"  Input: {repr(text)}")
    ws, rt = extract_prices(text, aggressive=True)
    print(f"  Result: WS=${ws}, RT=${rt}")
    print(f"  Expected: WS=${expected[0]}, RT=${expected[1]}")
    
    if ws == expected[0] and rt == expected[1]:
        print(f"  ✅ PASS")
    else:
        print(f"  ❌ FAIL")
        if ws != expected[0]:
            print(f"    - Wholesale mismatch: {ws} vs {expected[0]}")
        if rt != expected[1]:
            print(f"    - Retail mismatch: {rt} vs {expected[1]}")

print("\n" + "=" * 60)
