# Slide 255 Price Extraction Fix

## Issue Found
Slide 255 contains "NIKE AIR SPORT 2 ALP GB" golf bags with the format:
```
Wholesale: $143.00
Retail: $260.00
```

But extraction was showing both prices as **143**, not **143 and 260**.

## Root Cause
The regex pattern for extracting prices with labels before the value wasn't handling the **colon** (`:`) that appears after the label in this slide format.

**Pattern that failed:**
```regex
r"\b(WHOLESALE|WS|WHSE|RETAIL)\b\s*\$\s*(\d+(?:\.\d{2})?)"
```
This required the label immediately before `$`, but in slide 255:
```
Wholesale: $143.00  ← Has colon between label and $
```

## Solution Applied

### 1. Made colon optional in the regex
```python
# OLD: r"\b(WHOLESALE|WS|WHSE|RETAIL)\b\s*\$\s*(\d+(?:\.\d{2})?)"
# NEW: r"\b(WHOLESALE|WS|WHSE|RETAIL)\b\s*:?\s*\$\s*(\d+(?:\.\d{2})?)"
                                              ^^^ colon is now optional
```

Now handles both:
- `RETAIL $26.00` (no colon)
- `Retail: $260.00` (with colon)

### 2. Improved label extraction
Also refined the first pattern to be more robust:
```python
# Now uses the full matched text to determine label type
label_text = match.group(0).upper()
if any(label in label_text for label in ["WHOLESALE", "WS", "WHSE"]) and wholesale is None:
    wholesale = value
```

## Test Cases Verified

✅ **Format 1 - Apparel style** (no colon)
```
$70.00 WHOLESALE | $140.00 RETAIL
Result: Wholesale=70, Retail=140
```

✅ **Format 2 - Bag style with colon** (Slide 255)
```
Wholesale: $143.00
Retail: $260.00
Result: Wholesale=143, Retail=260
```

✅ **Format 3 - Mixed**
```
N.101.4771.320 Wholesale: $143.00 Retail: $260.00
Result: Wholesale=143, Retail=260
```

## Files Modified
- `extractor.py` - Enhanced regex patterns for price extraction
- `test_integration.py` - Added test case for Slide 255 format

## Before vs After

### Slide 255 Extraction Before Fix:
| Field | Value |
|-------|-------|
| Product | NIKE AIR SPORT 2 ALP GB |
| Style Code | N.101.4771 |
| Wholesale | 143 |
| Retail | **143** ❌ |

### Slide 255 Extraction After Fix:
| Field | Value |
|-------|-------|
| Product | NIKE AIR SPORT 2 ALP GB |
| Style Code | N.101.4771 |
| Wholesale | 143 |
| Retail | **260** ✅ |

---

**Status:** ✅ Fixed  
**Cause:** Regex pattern not accounting for colons after price labels  
**Solution:** Made colon optional in regex pattern  
**Impact:** Now handles both colon and non-colon formats
