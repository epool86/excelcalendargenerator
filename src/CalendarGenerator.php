<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;

class CalendarGenerator
{
    private const MONTHS = [
        1 => 'JANUARY',
        2 => 'FEBRUARY',
        3 => 'MARCH',
        4 => 'APRIL',
        5 => 'MAY',
        6 => 'JUNE',
        7 => 'JULY',
        8 => 'AUGUST',
        9 => 'SEPTEMBER',
        10 => 'OCTOBER',
        11 => 'NOVEMBER',
        12 => 'DECEMBER'
    ];

    private const DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];

    // Color constants
    private const COLOR_DARK_PURPLE = '5B2C6F';
    private const COLOR_LIGHT_PURPLE = 'AF7AC5';
    private const COLOR_WEEKEND_HEADER = '808080';
    private const COLOR_WEEKEND_CELL = 'D3D3D3';
    private const COLOR_MONTH_HEADER = '2E86AB';
    private const COLOR_YEAR_HEADER = '1B4F72';
    private const COLOR_HOLIDAY_TEXT = 'C0392B';
    private const COLOR_HOLIDAY_BG = 'FADBD8';

    // Row heights
    private const HEIGHT_YEAR_ROW = 35;
    private const HEIGHT_MONTH_ROW = 30;
    private const HEIGHT_DAY_HEADER_ROW = 25;
    private const HEIGHT_DATE_ROW = 20;
    private const HEIGHT_ACTIVITY_ROW = 50;

    // Column width (increased for more activity space)
    private const COLUMN_WIDTH = 22;

    private array $holidays;

    public function __construct()
    {
        $this->holidays = $this->loadHolidays();
    }

    private function loadHolidays(): array
    {
        return [
            // 2025 Holidays
            '2025-01-01' => 'New Year\'s Day',
            '2025-01-29' => 'Chinese New Year',
            '2025-03-31' => 'Hari Raya Aidilfitri',
            '2025-04-01' => 'Hari Raya Aidilfitri Holiday',
            '2025-05-01' => 'Labour Day',
            '2025-05-12' => 'Wesak Day',
            '2025-06-02' => 'Birthday of SPB Yang di Pertuan Agong',
            '2025-06-07' => 'Hari Raya Haji',
            '2025-06-27' => 'Awal Muharram',
            '2025-08-31' => 'National Day',
            '2025-09-05' => 'Maulidur Rasul',
            '2025-09-16' => 'Malaysia Day',
            '2025-12-25' => 'Christmas Day',

            // 2026 Holidays
            '2026-01-01' => 'New Year\'s Day',
            '2026-02-17' => 'Chinese New Year',
            '2026-02-18' => 'Chinese New Year Holiday',
            '2026-03-21' => 'Hari Raya Aidilfitri',
            '2026-03-22' => 'Hari Raya Aidilfitri Holiday',
            '2026-05-01' => 'Labour Day',
            '2026-05-27' => 'Hari Raya Haji',
            '2026-05-31' => 'Wesak Day',
            '2026-06-01' => 'Birthday of SPB Yang di Pertuan Agong',
            '2026-06-17' => 'Awal Muharram',
            '2026-08-25' => 'Maulidur Rasul',
            '2026-08-31' => 'National Day',
            '2026-09-16' => 'Malaysia Day',
            '2026-12-25' => 'Christmas Day',
        ];
    }

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
        $sheet = new Worksheet($spreadsheet, self::MONTHS[$month]);
        $spreadsheet->addSheet($sheet);

        // Set column widths
        foreach (range('A', 'G') as $col) {
            $sheet->getColumnDimension($col)->setWidth(self::COLUMN_WIDTH);
        }

        // Row 1: Year header
        $this->createYearHeader($sheet, $year);

        // Row 2: Month header
        $this->createMonthHeader($sheet, $month);

        // Row 3: Day of week headers
        $this->createDayHeaders($sheet);

        // Calendar grid
        $this->createCalendarGrid($sheet, $year, $month);
    }

    private function createYearHeader(Worksheet $sheet, int $year): void
    {
        $sheet->mergeCells('A1:G1');
        $sheet->setCellValue('A1', $year);
        $sheet->getRowDimension(1)->setRowHeight(self::HEIGHT_YEAR_ROW);

        $sheet->getStyle('A1:G1')->applyFromArray([
            'font' => [
                'bold' => true,
                'size' => 24,
                'color' => ['rgb' => 'FFFFFF'],
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => self::COLOR_YEAR_HEADER],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);
    }

    private function createMonthHeader(Worksheet $sheet, int $month): void
    {
        $sheet->mergeCells('A2:G2');
        $sheet->setCellValue('A2', self::MONTHS[$month]);
        $sheet->getRowDimension(2)->setRowHeight(self::HEIGHT_MONTH_ROW);

        $sheet->getStyle('A2:G2')->applyFromArray([
            'font' => [
                'bold' => true,
                'size' => 20,
                'color' => ['rgb' => 'FFFFFF'],
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => self::COLOR_MONTH_HEADER],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);
    }

    private function createDayHeaders(Worksheet $sheet): void
    {
        $sheet->getRowDimension(3)->setRowHeight(self::HEIGHT_DAY_HEADER_ROW);

        foreach (self::DAYS as $index => $day) {
            $col = chr(65 + $index); // A, B, C, D, E, F, G
            $sheet->setCellValue($col . '3', $day);

            // Determine background color
            if ($index >= 5) {
                // Weekend (Sat, Sun)
                $bgColor = self::COLOR_WEEKEND_HEADER;
            } else {
                // Striped pattern: dark, light, dark, light, dark
                $bgColor = ($index % 2 === 0) ? self::COLOR_DARK_PURPLE : self::COLOR_LIGHT_PURPLE;
            }

            $sheet->getStyle($col . '3')->applyFromArray([
                'font' => [
                    'bold' => true,
                    'size' => 14,
                    'color' => ['rgb' => 'FFFFFF'],
                ],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['rgb' => $bgColor],
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_CENTER,
                ],
            ]);
        }
    }

    private function createCalendarGrid(Worksheet $sheet, int $year, int $month): void
    {
        $daysInMonth = cal_days_in_month(CAL_GREGORIAN, $month, $year);
        $firstDayOfWeek = $this->getFirstDayOfWeek($year, $month);

        $currentRow = 4;
        $currentDay = 1;
        $dayOfWeek = $firstDayOfWeek;

        // Calculate number of weeks needed
        $totalCells = $firstDayOfWeek + $daysInMonth;
        $totalWeeks = ceil($totalCells / 7);

        for ($week = 0; $week < $totalWeeks; $week++) {
            $dateRow = $currentRow;
            $activityRow = $currentRow + 1;

            // Set row heights
            $sheet->getRowDimension($dateRow)->setRowHeight(self::HEIGHT_DATE_ROW);
            $sheet->getRowDimension($activityRow)->setRowHeight(self::HEIGHT_ACTIVITY_ROW);

            for ($col = 0; $col < 7; $col++) {
                $colLetter = chr(65 + $col);
                $dateCell = $colLetter . $dateRow;
                $activityCell = $colLetter . $activityRow;

                // Check if we should display a date
                $showDate = false;
                $dateNumber = null;

                if ($week === 0 && $col < $firstDayOfWeek) {
                    // Before the first day of the month
                    $showDate = false;
                } elseif ($currentDay <= $daysInMonth) {
                    if ($week === 0 && $col >= $firstDayOfWeek) {
                        $showDate = true;
                        $dateNumber = $currentDay;
                        $currentDay++;
                    } elseif ($week > 0) {
                        $showDate = true;
                        $dateNumber = $currentDay;
                        $currentDay++;
                    }
                }

                // Determine if weekend
                $isWeekend = ($col >= 5);

                // Check for holiday
                $holidayName = null;
                $isHoliday = false;
                if ($showDate && $dateNumber !== null) {
                    $dateKey = sprintf('%04d-%02d-%02d', $year, $month, $dateNumber);
                    if (isset($this->holidays[$dateKey])) {
                        $holidayName = $this->holidays[$dateKey];
                        $isHoliday = true;
                    }
                }

                // Determine background color (holiday takes priority, then weekend, then white)
                if ($isHoliday) {
                    $bgColor = self::COLOR_HOLIDAY_BG;
                } elseif ($isWeekend) {
                    $bgColor = self::COLOR_WEEKEND_CELL;
                } else {
                    $bgColor = 'FFFFFF';
                }

                // Style date cell - no bottom border (merged look with activity row)
                $dateStyle = [
                    'font' => [
                        'size' => 11,
                        'bold' => true,
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_RIGHT,
                        'vertical' => Alignment::VERTICAL_TOP,
                    ],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => $bgColor],
                    ],
                    'borders' => [
                        'top' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['rgb' => '000000'],
                        ],
                        'left' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['rgb' => '000000'],
                        ],
                        'right' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['rgb' => '000000'],
                        ],
                        // No bottom border - seamless with activity row
                    ],
                ];

                // Style activity cell - no top border (merged look with date row)
                $activityStyle = [
                    'font' => [
                        'size' => 10,
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_LEFT,
                        'vertical' => Alignment::VERTICAL_TOP,
                        'wrapText' => true,
                    ],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => $bgColor],
                    ],
                    'borders' => [
                        'bottom' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['rgb' => '000000'],
                        ],
                        'left' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['rgb' => '000000'],
                        ],
                        'right' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['rgb' => '000000'],
                        ],
                        // No top border - seamless with date row
                    ],
                ];

                // Apply holiday text styling
                if ($isHoliday) {
                    $activityStyle['font']['bold'] = true;
                    $activityStyle['font']['color'] = ['rgb' => self::COLOR_HOLIDAY_TEXT];
                }

                // Set cell values
                if ($showDate && $dateNumber !== null) {
                    $sheet->setCellValue($dateCell, $dateNumber);
                    if ($holidayName) {
                        $sheet->setCellValue($activityCell, $holidayName);
                    }
                }

                // Apply styles
                $sheet->getStyle($dateCell)->applyFromArray($dateStyle);
                $sheet->getStyle($activityCell)->applyFromArray($activityStyle);
            }

            $currentRow += 2;
        }

        // Apply thick outer border to entire calendar
        $lastRow = $currentRow - 1;
        $sheet->getStyle("A1:G{$lastRow}")->applyFromArray([
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);

        // Thick border below headers
        $sheet->getStyle('A3:G3')->applyFromArray([
            'borders' => [
                'bottom' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);
    }

    /**
     * Get the day of week for the first day of the month (0 = Monday, 6 = Sunday)
     */
    private function getFirstDayOfWeek(int $year, int $month): int
    {
        $date = new \DateTime("{$year}-{$month}-01");
        $dayOfWeek = (int) $date->format('N'); // 1 (Monday) to 7 (Sunday)
        return $dayOfWeek - 1; // Convert to 0-indexed (0 = Monday)
    }
}
