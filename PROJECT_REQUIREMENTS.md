# PHP Excel Calendar Generator - Project Requirements & Plan

## Overview

A simple PHP web application that generates downloadable Excel files containing a 12-month calendar for a selected year. Each month is displayed on a separate sheet with a horse-race style calendar layout, allowing users to write activities for each day.

---

## Technology Stack

### Recommended Package: **PhpSpreadsheet**

**Why PhpSpreadsheet?**
- Most comprehensive Excel manipulation library for PHP
- Full support for `.xlsx` format (Excel 2007+)
- Extensive formatting capabilities:
  - Cell background colors
  - Borders (thin, thick, colored)
  - Cell merging
  - Font styling (bold, size, color)
  - Row/column sizing
  - Multiple worksheets
- Active maintenance (by PHPOffice)
- MIT License
- Composer installation: `composer require phpoffice/phpspreadsheet`

---

## Functional Requirements

### 1. Web Interface
- Simple HTML page with:
  - Year dropdown selector (2025, 2026, 2027, etc.)
  - "Generate & Download" button
- On button click: generate Excel file and trigger browser download

### 2. Excel File Structure
- **File format:** `.xlsx`
- **Filename:** `Calendar_{YEAR}.xlsx` (e.g., `Calendar_2025.xlsx`)
- **Sheets:** 12 sheets, one per month (January to December)

### 3. Sheet Layout (Per Month)

Each sheet follows this structure from top to bottom:

```
+------------------------------------------------------------------+
|                           2025                                    |  <- Row 1: Year (merged, large font, centered)
+------------------------------------------------------------------+
|                         JANUARY                                   |  <- Row 2: Month name (merged, large font, bg color)
+------------------------------------------------------------------+
|  MON  |  TUE  |  WED  |  THU  |  FRI  |  SAT  |  SUN  |          |  <- Row 3: Day headers (bold, striped bg)
+-------+-------+-------+-------+-------+-------+-------+          |
|   1   |   2   |   3   |   4   |   5   |   6   |   7   |          |  <- Date row (small, right-aligned)
|       |       |       |       |       |  grey |  grey |          |  <- Activity row (tall, for user input)
+-------+-------+-------+-------+-------+-------+-------+          |
|   8   |   9   |  10   |  11   |  12   |  13   |  14   |          |
|       |       |       |       |       |  grey |  grey |          |
+-------+-------+-------+-------+-------+-------+-------+          |
|  ...continues for all weeks of the month...                      |
+------------------------------------------------------------------+
```

### 4. Styling Specifications

#### Header Section (Rows 1-2)
| Element | Style |
|---------|-------|
| Year | Font size: 24pt, Bold, Centered, Merged across all columns |
| Month Name | Font size: 20pt, Bold, Centered, Merged across all columns, Background color (unique per month or consistent) |

#### Day-of-Week Headers (Row 3)
| Element | Style |
|---------|-------|
| Font | Size: 14pt, Bold, Centered |
| Background | Alternating/striped: Dark Purple (#5B2C6F) → Light Purple (#AF7AC5) → Dark Purple... |
| Saturday & Sunday | Grey background (#808080) instead of purple |

#### Calendar Days (Repeating 2-row pattern per week)
| Row Type | Style |
|----------|-------|
| **Date Row** | Height: ~20px, Font size: 11pt, Right-aligned, Top-aligned |
| **Activity Row** | Height: ~60px (tall for user input), Font size: 10pt, Top-left aligned |
| **Weekend Columns (Sat/Sun)** | Grey background (#D3D3D3) for both date and activity rows |
| **Weekday Columns** | White/light background |

#### Borders
| Area | Border Style |
|------|--------------|
| Outer calendar border | Thick black border (medium or thick weight) |
| Inner cell borders | Thin black borders |
| Header section | Thick bottom border separating from calendar grid |

#### Column Widths
- Each day column: ~15-18 characters wide (to allow activity text)
- Total 7 columns for Mon-Sun

---

## Technical Implementation Plan

### Phase 1: Project Setup
1. Initialize project directory structure
2. Create `composer.json`
3. Install PhpSpreadsheet via Composer
4. Create basic folder structure:
   ```
   /excel
   ├── composer.json
   ├── vendor/
   ├── public/
   │   └── index.php          # Web interface
   ├── src/
   │   └── CalendarGenerator.php   # Excel generation logic
   └── PROJECT_REQUIREMENTS.md
   ```

### Phase 2: Core Calendar Logic
1. Create `CalendarGenerator` class with methods:
   - `generate(int $year): Spreadsheet` - Main generation method
   - `createMonthSheet(Spreadsheet $spreadsheet, int $year, int $month)` - Creates single month sheet
   - `getMonthStartDay(int $year, int $month): int` - Get day of week for 1st of month
   - `getDaysInMonth(int $year, int $month): int` - Get total days in month
   - `applyHeaderStyles(Worksheet $sheet)` - Style year/month headers
   - `applyCalendarStyles(Worksheet $sheet, int $totalWeeks)` - Style calendar grid
   - `applyBorders(Worksheet $sheet, string $range)` - Apply border formatting

### Phase 3: Excel Formatting Implementation
1. Define color constants for:
   - Dark purple (header stripe)
   - Light purple (header stripe alternate)
   - Grey (weekends)
   - Month header background
2. Implement header row creation with merged cells
3. Implement day-of-week header row with striped backgrounds
4. Implement calendar grid with:
   - Proper date placement based on starting day of month
   - 2-row per week pattern (date row + activity row)
   - Weekend column highlighting
5. Apply thick outer borders and thin inner borders
6. Set row heights and column widths

### Phase 4: Web Interface
1. Create simple HTML form in `public/index.php`:
   - Year dropdown (dynamically generated, current year + next 5 years)
   - Submit button styled nicely
2. Handle form submission:
   - Validate year input
   - Call `CalendarGenerator`
   - Set proper HTTP headers for Excel download
   - Output file and exit

### Phase 5: Testing & Refinement
1. Test generated Excel in:
   - Microsoft Excel (Windows)
   - LibreOffice Calc (Linux)
   - Google Sheets (import)
2. Verify all formatting renders correctly
3. Test leap year handling (February 29)
4. Test months starting on different days
5. Adjust styling as needed

---

## File Structure (Final)

```
/excel
├── composer.json
├── composer.lock
├── vendor/                     # Composer dependencies
├── public/
│   └── index.php              # Entry point & web UI
├── src/
│   └── CalendarGenerator.php  # Core generation class
├── PROJECT_REQUIREMENTS.md    # This document
└── .claude/                   # Claude config (ignore)
```

---

## Color Palette

| Purpose | Color Name | Hex Code |
|---------|------------|----------|
| Dark Purple (header stripe) | Rebecca Purple | `#5B2C6F` |
| Light Purple (header stripe) | Medium Purple | `#AF7AC5` |
| Weekend Background | Light Grey | `#D3D3D3` |
| Weekend Header | Dark Grey | `#808080` |
| Month Header | Deep Sky Blue | `#2E86AB` |
| Year Header | Navy | `#1B4F72` |
| Borders | Black | `#000000` |

---

## Sample Code Structure

```php
<?php
// src/CalendarGenerator.php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class CalendarGenerator
{
    private const MONTHS = [
        1 => 'January', 2 => 'February', 3 => 'March',
        4 => 'April', 5 => 'May', 6 => 'June',
        7 => 'July', 8 => 'August', 9 => 'September',
        10 => 'October', 11 => 'November', 12 => 'December'
    ];

    private const DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];

    public function generate(int $year): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();

        // Remove default sheet
        $spreadsheet->removeSheetByIndex(0);

        // Create 12 month sheets
        for ($month = 1; $month <= 12; $month++) {
            $this->createMonthSheet($spreadsheet, $year, $month);
        }

        // Set first sheet as active
        $spreadsheet->setActiveSheetIndex(0);

        return $spreadsheet;
    }

    private function createMonthSheet(Spreadsheet $spreadsheet, int $year, int $month): void
    {
        // Implementation here...
    }
}
```

---

## API/Download Endpoint

```php
<?php
// public/index.php

require_once __DIR__ . '/../vendor/autoload.php';

use App\CalendarGenerator;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['year'])) {
    $year = (int) $_POST['year'];

    // Validate year
    if ($year < 2020 || $year > 2100) {
        die('Invalid year');
    }

    // Generate Excel
    $generator = new CalendarGenerator();
    $spreadsheet = $generator->generate($year);

    // Set headers for download
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="Calendar_' . $year . '.xlsx"');
    header('Cache-Control: max-age=0');

    // Output file
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    exit;
}
?>
<!-- HTML form here -->
```

---

## Success Criteria

- [ ] User can select a year from dropdown
- [ ] Clicking "Generate" downloads an `.xlsx` file
- [ ] Excel file contains 12 sheets (one per month)
- [ ] Each sheet has year and month headers with proper styling
- [ ] Day-of-week headers have alternating purple stripe pattern
- [ ] Calendar displays correct dates for each month
- [ ] Each day has 2 rows (date number + activity space)
- [ ] Activity rows are tall enough for user input
- [ ] Saturday and Sunday columns have grey background
- [ ] Thick borders around entire calendar
- [ ] Thin borders between cells
- [ ] File opens correctly in Excel/LibreOffice/Google Sheets
- [ ] Leap years handled correctly (Feb 29)

---

## Malaysian Public Holidays Data

The calendar will automatically populate Malaysian national public holidays in the activity cells.

### 2025 Holidays

| Date | Holiday Name |
|------|--------------|
| January 1 | New Year's Day |
| January 29 | Chinese New Year |
| March 31 | Hari Raya Aidilfitri |
| April 1 | Hari Raya Aidilfitri Holiday |
| May 1 | Labour Day |
| May 12 | Wesak Day |
| June 2 | Birthday of SPB Yang di Pertuan Agong |
| June 7 | Hari Raya Haji |
| June 27 | Awal Muharram |
| August 31 | National Day |
| September 5 | Maulidur Rasul |
| September 16 | Malaysia Day |
| December 25 | Christmas Day |

### 2026 Holidays

| Date | Holiday Name |
|------|--------------|
| January 1 | New Year's Day |
| February 17 | Chinese New Year |
| February 18 | Chinese New Year Holiday |
| March 21 | Hari Raya Aidilfitri |
| March 22 | Hari Raya Aidilfitri Holiday |
| May 1 | Labour Day |
| May 27 | Hari Raya Haji |
| May 31 | Wesak Day |
| June 1 | Birthday of SPB Yang di Pertuan Agong |
| June 17 | Awal Muharram |
| August 25 | Maulidur Rasul |
| August 31 | National Day |
| September 16 | Malaysia Day |
| December 25 | Christmas Day |

### Holiday Cell Styling

| Element | Style |
|---------|-------|
| Holiday Text | Red font color (`#C0392B`), Bold |
| Holiday Background | Light red/pink (`#FADBD8`) |
| Cell Content | Holiday name displayed in the activity row |

### Implementation Notes

- Holidays will be stored as a PHP array keyed by `YYYY-MM-DD` format
- The generator will look up each date and populate the activity cell if a holiday exists
- Holiday cells will have distinctive styling (red text, light pink background) to stand out
- Users can still type additional activities alongside the holiday text

---

## Notes

- Week starts on **Monday** (European/ISO standard)
- Empty cells before the 1st of month and after last day should remain blank but styled
- Consider adding print area settings for better printing
- Holiday data source: [Office Holidays Malaysia](https://www.officeholidays.com/countries/malaysia)
