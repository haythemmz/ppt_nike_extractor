# Nike PPT Extractor - Bag/Accessories Integration

## Summary
Successfully integrated bag/accessories extraction functionality from `extractor_with_bags.py` into the main `extractor.py` module with **enhanced price extraction for bags**.

## Changes Made

### 1. **Updated Regex Patterns** (`extractor.py`)
   - Added `ACCESSORY_STYLE_COLOR_RE` to match Nike bag/accessory codes like `N.100.3478.091`
   - Updated `STOP_RE` to include "Material Content" and "Features and Benefits" for better section extraction

### 2. **New Helper Functions** (`extractor.py`)
   - `split_accessory_style_color()` - Splits bag codes like `N.100.3478.091` into style code (`N.100.3478`) and color code (`091`)
   - `has_accessory_style_color()` - Checks if text contains accessory/bag codes
   - `build_style_color_key()` - Builds the proper style-color key format, handling both:
     - Accessories: `N.100.3478.091` format
     - Apparel: `STYLE-COLOR` format

### 3. **Enhanced Price Extraction** 🆕
   - Added `aggressive` parameter to `extract_prices()` function
   - Aggressive mode intelligently handles bags where prices may lack clear WHOLESALE/RETAIL labels
   - When aggressive=True:
     - Looks at up to 10 lines (instead of 5) to find both prices
     - Automatically assigns smaller price to wholesale and larger to retail
     - Works even when price labels are unclear or on separate lines

### 4. **Special Bag Handling in Main Extraction** (`extractor.py`)
   - Added dedicated loop before regular product extraction to identify and process bag slides
   - Uses aggressive price extraction mode for bags
   - Extracts material content and key features for bags
   - Properly parses bag prices and status information
   - Marks processed lines to avoid double-processing

### 5. **Updated Style-Color Key Building**
   - Changed from hardcoded `f"{style}-{color}"` to use `build_style_color_key()` function
   - Ensures bags get proper dot notation format while apparel retains hyphen format

### 6. **Updated UI Description** (`app.py`)
   - Added mention of accessories/bags support in the app description
   - Users now see that the tool supports both apparel and Nike accessories

## Features Now Supported

✅ Regular Nike apparel products (shirts, pants, shoes, etc.)
✅ Nike accessories and bags with special style codes (N.XXX.XXXX.XXX format)
✅ Proper color code extraction for both product types
✅ **Fixed price extraction for bags** - Now correctly finds wholesale AND retail prices
✅ Price extraction works even when prices are on separate lines
✅ Material and features extraction for accessories
✅ Status tracking (NEW/CARRYOVER) for all products
✅ Separate style-color key formats by product type

## File Changes
- `extractor.py` - Enhanced with improved bag extraction and price handling
- `app.py` - Updated UI description
- `test_integration.py` - Added price extraction tests
- `extractor_with_bags.py` - Can now be archived or removed (functionality merged into extractor.py)

## Testing
The updated app will automatically handle PowerPoint files containing:
- Pure apparel products
- Pure accessory/bag products  
- Mixed slides with both apparel and accessories

All three output Excel sheets (Products, Style_Colors, Raw_Slide_Text) remain compatible with previous versions.

### Key Fix for Bag Prices
**Before:** Bags showed same wholesale and retail prices  
**After:** Both prices are correctly extracted and differentiated, even when:
- Prices appear on separate lines
- Price labels are unclear
- Prices are on multiple lines after the bag code

