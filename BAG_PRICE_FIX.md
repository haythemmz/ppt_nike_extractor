# Bag Price Extraction Fix

## Problem
Bags were showing the same wholesale and retail prices instead of different values.

**Example:**
```
Extracted Data:
- Wholesale: $70.00
- Retail: $70.00  ❌ Should be $140.00
```

## Root Causes
1. **Limited line context**: Only combining 5 lines, but prices might be on separate lines further down
2. **No aggressive price search**: When prices lacked clear labels, extraction would fail or get stuck on first price
3. **No price comparison logic**: Didn't intelligently assign smaller vs larger price to wholesale vs retail

## Solution Implemented

### 1. **Extended Line Context**
```python
# Before: Only 5 lines
combined = clean_text(" ".join(lines[i : min(i + 5, len(lines))]))

# After: 10 lines for bags
combined = clean_text(" ".join(lines[i : min(i + 10, len(lines))]))
```

### 2. **Aggressive Price Extraction Mode**
Added new parameter to `extract_prices()`:
```python
def extract_prices(text: str, aggressive: bool = False):
    # ... existing logic ...
    
    # New: If aggressive mode and we have 2+ prices but retail is missing
    if aggressive and wholesale is not None and retail is None and len(values) >= 2:
        sorted_values = sorted(set(values))
        if sorted_values[0] == wholesale and len(sorted_values) > 1:
            retail = sorted_values[1]  # Use second (larger) price
```

### 3. **Use Aggressive Mode for Bags**
```python
# Bag extraction uses aggressive=True
wholesale, retail = extract_prices(combined, aggressive=True)

# Regular apparel still uses default mode
wholesale, retail = extract_prices(combined)  # aggressive=False by default
```

## Result
✅ Bags now correctly extract both wholesale AND retail prices  
✅ Works even when prices are on separate lines  
✅ Handles cases where labels are missing or unclear  
✅ Regular apparel extraction unchanged (backward compatible)

## Example - Before vs After

### PowerPoint Slide (Bags section):
```
N.100.3478.091
Material: Polyester
Key Features: Durable
$70.00 WHOLESALE
$140.00 RETAIL
NEW
```

### Before Fix:
| Field | Value |
|-------|-------|
| Style Code | N.100.3478 |
| Color Code | 091 |
| Wholesale | $70.00 |
| Retail | $70.00 ❌ |

### After Fix:
| Field | Value |
|-------|-------|
| Style Code | N.100.3478 |
| Color Code | 091 |
| Wholesale | $70.00 ✅ |
| Retail | $140.00 ✅ |
