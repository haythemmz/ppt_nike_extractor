#!/usr/bin/env python
"""Quick test to verify bag extraction integration"""

# Test that the updated extractor module loads correctly
try:
    from extractor import (
        ACCESSORY_STYLE_COLOR_RE,
        split_accessory_style_color,
        has_accessory_style_color,
        build_style_color_key,
        build_image_name,
        expand_standard_style_color_tokens,
        is_valid_style_color_key,
        parse_detail_style_line,
        extract_colors_from_style_color_tokens,
        detect_slide_category,
        accessory_slide_product_name,
        get_accessory_price_context,
        extract_prices,
        extract_products_from_pptx,
    )
    print("✓ All bag extraction functions imported successfully")
    
    # Test regex
    test_code = "N.100.3478.091"
    if ACCESSORY_STYLE_COLOR_RE.search(test_code):
        print(f"✓ Regex correctly matches: {test_code}")
    else:
        print(f"✗ Regex failed to match: {test_code}")
    
    # Test split function
    style, color = split_accessory_style_color(test_code)
    assert style == "N.100.3478", f"Style mismatch: {style}"
    assert color == "091", f"Color mismatch: {color}"
    print(f"✓ split_accessory_style_color works: {test_code} → {style}.{color}")
    
    # Test style-color key for bags
    key_bag = build_style_color_key("N.100.3478", "091")
    assert key_bag == "N.100.3478.091", f"Bag key mismatch: {key_bag}"
    print(f"✓ Bag style-color key: {key_bag}")
    
    # Test style-color key for apparel
    key_apparel = build_style_color_key("IR4565", "010")
    assert key_apparel == "IR4565-010", f"Apparel key mismatch: {key_apparel}"
    print(f"✓ Apparel style-color key: {key_apparel}")

    image_name = build_image_name(12, "IR4565-010")
    assert image_name == "slide_012_IR4565-010.png", f"Image name mismatch: {image_name}"
    print(f"✓ Image name: {image_name}")

    colors = extract_colors_from_style_color_tokens("W NK DF VLCITY SL POLO IO1640-100 $30.00 WS | $60.00 RETAIL")
    assert colors == ["100"], f"Style-color token color extraction failed: {colors}"
    print(f"✓ Grid style-color token color: {colors[0]}")

    slash_codes = expand_standard_style_color_tokens("IR5946-010/676")
    slash_keys = [row["style_color_key"] for row in slash_codes]
    assert slash_keys == ["IR5946-010", "IR5946-676"], f"Slash color expansion failed: {slash_keys}"
    print(f"Slash color expansion: {slash_keys}")

    assert is_valid_style_color_key("IR5946-010"), "Standard style-color validation failed"
    assert is_valid_style_color_key("N.101.4771.320"), "Accessory style-color validation failed"
    assert not is_valid_style_color_key("RETAIL"), "Invalid style-color validation failed"
    assert not is_valid_style_color_key("RETAIL-010"), "Fake RETAIL style-color validation failed"
    print("Strict style-color validation works")

    detail = parse_detail_style_line("IZ4772 | $65.00 WHOLESALE | $130.00 RETAIL | NEW")
    assert detail and detail["style"] == "IZ4772", f"Detail style line failed: {detail}"
    assert parse_detail_style_line("RETAIL | $65.00 WHOLESALE | $130.00 RETAIL | NEW") is None
    print("Strict detail style line parsing works")

    from extractor import is_likely_product_name
    assert is_likely_product_name("W NK AF TP SS POLO"), "Real apparel product name was rejected"
    assert not is_likely_product_name("Top stitched self-collar"), "Feature text accepted as product name"
    assert not is_likely_product_name("Utility webbing conveniently holds your accessories"), "Bag feature text accepted as product name"
    assert not is_likely_product_name("90% Polyester"), "Material text accepted as product name"
    assert not is_likely_product_name("NIKE Global Sports Apparel"), "Deck header accepted as product name"
    assert is_likely_product_name("NIKE AIR HYBRID 2 GB"), "Nike bag product name was rejected"
    print("Feature/material text is rejected as product name")

    category = detect_slide_category(["NIKE Global Sports Apparel", "WOMEN'S APPAREL", "Keep it Tight."])
    assert category == "Women's Apparel", f"Category detection failed: {category}"
    print(f"✓ Category detection: {category}")
    bag_category = detect_slide_category(["NIKE Global Sports Apparel", "BAGS", "Keep it Tight."])
    assert bag_category == "Bags", f"Bag category detection failed: {bag_category}"
    print(f"Bag category detection: {bag_category}")

    false_bag_category = detect_slide_category(["W NK DF PANT DRAWSTRING", "IR4565 | $70.00 WHOLESALE | $140.00 RETAIL"])
    assert false_bag_category is None, f"Product name misclassified as category: {false_bag_category}"
    print("Drawstring pant product name does not reset category")

    primary_bag_name = accessory_slide_product_name([
        "NIKE Global Sports Apparel",
        "BAGS",
        "STANDBAGS",
        "NIKE AIR SPORT 2 ALP GB",
        "NIKE AIR SPORT 2 ALP GB",
        "N.101.4771.320",
        "NIKE AIR HYBRID 2 GB",
        "N.100.3478.091",
    ])
    assert primary_bag_name == "NIKE AIR SPORT 2 ALP GB", f"Primary bag title failed: {primary_bag_name}"
    print(f"Primary bag title: {primary_bag_name}")

    # Test price extraction with aggressive mode
    print("\n--- Testing price extraction ---")
    
    # Normal case with labels
    ws, rt = extract_prices("$70 WHOLESALE $140 RETAIL")
    assert ws == 70.0 and rt == 140.0, f"Normal price extraction failed: {ws}, {rt}"
    print(f"✓ Normal prices: ${ws} WS, ${rt} RT")
    
    # Aggressive mode - two prices without labels
    ws, rt = extract_prices("$70.00 $140.00", aggressive=False)
    assert ws == 70.0 and rt == 140.0, f"Fallback price extraction failed: {ws}, {rt}"
    print(f"✓ Fallback prices: ${ws} WS, ${rt} RT")
    
    # Test aggressive mode with only wholesale label
    ws, rt = extract_prices("$70.00 WHOLESALE $140.00", aggressive=True)
    assert ws == 70.0 and rt == 140.0, f"Aggressive price extraction failed: {ws}, {rt}"
    print(f"✓ Aggressive mode prices: ${ws} WS, ${rt} RT")
    
    # Test on multiple lines (like in bags)
    bag_text = "N.100.3478.091\n70.00 WHOLESALE\n140.00 RETAIL"
    ws, rt = extract_prices(bag_text, aggressive=True)
    assert ws == 70.0 and rt == 140.0, f"Multi-line bag prices failed: {ws}, {rt}"
    print(f"✓ Bag prices from multi-line: ${ws} WS, ${rt} RT")
    
    # NEW: Test with just spaces between prices (no / separator)
    bag_space = "Wholesale: $143.00 Retail: $260.00"
    ws, rt = extract_prices(bag_space, aggressive=True)
    assert ws == 143.0 and rt == 260.0, f"Space-separated prices failed: {ws}, {rt}"
    print(f"✓ Space-separated prices: ${ws} WS, ${rt} RT")
    
    # Test with / separator (gloves format)
    glove_text = "Wholesale: $13.50 / Retail: $27.00"
    ws, rt = extract_prices(glove_text, aggressive=True)
    assert ws == 13.50 and rt == 27.00, f"Slash-separated prices failed: {ws}, {rt}"
    print(f"✓ Slash-separated prices: ${ws} WS, ${rt} RT")
    
    # NEW: Test the actual format from slide 255: "Wholesale: $143.00 Retail: $260.00"
    slide_255_format = "Wholesale: $143.00 Retail: $260.00"
    ws, rt = extract_prices(slide_255_format, aggressive=True)
    assert ws == 143.0, f"Slide 255 wholesale failed: {ws}"
    assert rt == 260.0, f"Slide 255 retail failed: {rt}"
    print(f"✓ Slide 255 format: ${ws} WS, ${rt} RT")

    raw_slide_255_line = "NIKE AIR SPORT 2 ALP GB N.101.4771.320 Wholesale: $143.00 Retail: $260.00"
    ws, rt = extract_prices(raw_slide_255_line, aggressive=True)
    assert ws == 143.0, f"Slide 255 raw line wholesale failed: {ws}"
    assert rt == 260.0, f"Slide 255 raw line retail failed: {rt}"
    print(f"✓ Slide 255 raw line: ${ws} WS, ${rt} RT")

    # Retail-only lines should not get copied into wholesale.
    ws, rt = extract_prices("Retail: $260.00", aggressive=False)
    assert ws is None and rt == 260.0, f"Retail-only price copied to wholesale: {ws}, {rt}"
    print(f"✓ Retail-only price stays retail-only: WS={ws}, RT=${rt}")

    ws, rt = extract_prices("$260.00 RETAIL", aggressive=False)
    assert ws is None and rt == 260.0, f"Retail-after-price copied to wholesale: {ws}, {rt}"
    print(f"✓ Retail label after price stays retail-only: WS={ws}, RT=${rt}")

    ws, rt = extract_prices("Retail: 260.00", aggressive=False)
    assert ws is None and rt == 260.0, f"Retail-only no-dollar price copied to wholesale: {ws}, {rt}"
    print(f"✓ Retail-only no-dollar price stays retail-only: WS={ws}, RT=${rt}")

    ws, rt = extract_prices("143.00 Wholesale 260.00 Retail", aggressive=True)
    assert ws == 143.0 and rt == 260.0, f"No-dollar labeled prices failed: {ws}, {rt}"
    print(f"✓ No-dollar labeled prices: ${ws} WS, ${rt} RT")

    # Headcover/accessory captions should use the same text box, not nearby captions.
    fake_items = [
        {"shape_index": 1, "line_index": 0, "text": "NIKE ICON BLADE PUTTER", "left": 100, "top": 100, "width": 250, "height": 80},
        {"shape_index": 1, "line_index": 1, "text": "HEADCOVER", "left": 100, "top": 100, "width": 250, "height": 80},
        {"shape_index": 1, "line_index": 2, "text": "N.101.4761.016", "left": 100, "top": 100, "width": 250, "height": 80},
        {"shape_index": 1, "line_index": 3, "text": "Wholesale: $17.50", "left": 100, "top": 100, "width": 250, "height": 80},
        {"shape_index": 1, "line_index": 4, "text": "Retail: $35.00", "left": 100, "top": 100, "width": 250, "height": 80},
        {"shape_index": 2, "line_index": 0, "text": "N.101.4761.074", "left": 500, "top": 100, "width": 250, "height": 80},
        {"shape_index": 2, "line_index": 1, "text": "Retail: $99.00", "left": 500, "top": 100, "width": 250, "height": 80},
    ]
    context = get_accessory_price_context(fake_items, fake_items[2], [item["text"] for item in fake_items], 2)
    ws, rt = extract_prices(context, aggressive=True)
    assert ws == 17.5 and rt == 35.0, f"Accessory caption context mixed prices: {ws}, {rt}"
    print(f"✓ Accessory caption prices stay local: ${ws} WS, ${rt} RT")

    # Test with the actual bag code format
    bag_full = "N.101.4771.320 Wholesale: $143.00 Retail: $260.00"
    ws, rt = extract_prices(bag_full, aggressive=True)
    assert ws == 143.0, f"Full bag code wholesale failed: {ws}"
    assert rt == 260.0, f"Full bag code retail failed: {rt}"
    print(f"✓ Full bag format: ${ws} WS, ${rt} RT")
    
    print("\n✓ All integration tests passed!")
    
except ImportError as e:
    print(f"✗ Import failed: {e}")
    exit(1)
except AssertionError as e:
    print(f"✗ Test failed: {e}")
    exit(1)
except Exception as e:
    print(f"✗ Unexpected error: {e}")
    exit(1)

