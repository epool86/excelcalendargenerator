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
    private const COLOR_SCHOOL_HOLIDAY_BG = 'FFF9C4'; // Yellow for school holidays
    private const COLOR_SCHOOL_HOLIDAY_WEEKEND_BG = 'E6D98C'; // Darker yellow-gray for school holiday on weekend
    private const COLOR_HOLIDAY_SCHOOL_BG = 'FFEB3B'; // Bright yellow for public holiday + school holiday overlap

    // Row heights
    private const HEIGHT_YEAR_ROW = 35;
    private const HEIGHT_MONTH_ROW = 30;
    private const HEIGHT_DAY_HEADER_ROW = 25;
    private const HEIGHT_DATE_ROW = 20;
    private const HEIGHT_ACTIVITY_ROW = 50;

    // Column width (increased for more activity space)
    private const COLUMN_WIDTH = 22;

    private array $holidays;
    private array $schoolHolidays;

    public function __construct()
    {
        $this->holidays = $this->loadHolidays();
        $this->schoolHolidays = $this->loadSchoolHolidays();
    }

    private function loadHolidays(): array
    {
        return [
            // ============ 2025 National Holidays ============
            '2025-01-01' => 'New Year\'s Day',
            '2025-01-29' => 'Chinese New Year',
            '2025-01-30' => 'Chinese New Year (Day 2)',
            '2025-02-01' => 'Federal Territory Day',
            '2025-02-11' => 'Thaipusam',
            '2025-03-02' => 'Awal Ramadan',
            '2025-03-18' => 'Nuzul Al-Quran',
            '2025-03-31' => 'Hari Raya Aidilfitri',
            '2025-04-01' => 'Hari Raya Aidilfitri (Day 2)',
            '2025-05-01' => 'Labour Day',
            '2025-05-12' => 'Wesak Day',
            '2025-05-30' => 'Pesta Kaamatan (Sabah)',
            '2025-05-31' => 'Pesta Kaamatan (Sabah)',
            '2025-06-01' => 'Gawai Dayak (Sarawak)',
            '2025-06-02' => 'Gawai Dayak / Agong Birthday',
            '2025-06-07' => 'Hari Raya Haji',
            '2025-06-08' => 'Hari Raya Haji (Day 2)',
            '2025-06-27' => 'Awal Muharram',
            '2025-07-22' => 'Sarawak Day',
            '2025-08-31' => 'National Day',
            '2025-09-05' => 'Maulidur Rasul',
            '2025-09-16' => 'Malaysia Day',
            '2025-10-20' => 'Deepavali',
            '2025-12-25' => 'Christmas Day',

            // 2025 State Holidays
            '2025-01-14' => 'Hari Hol Negeri Sembilan',
            '2025-01-19' => 'Sultan Perak Birthday',
            '2025-02-05' => 'Sultan Kedah Birthday',
            '2025-03-04' => 'Israk Mikraj',
            '2025-03-23' => 'Sultan Johor Birthday',
            '2025-04-15' => 'Sultan Terengganu Birthday',
            '2025-04-19' => 'Sultan Perak Declaration',
            '2025-04-26' => 'Sultan Kelantan Birthday',
            '2025-05-17' => 'Raja Perlis Birthday',
            '2025-05-22' => 'Hari Hol Pahang',
            '2025-07-11' => 'Penang Governor Birthday',
            '2025-07-22' => 'Sultan Pahang Birthday',
            '2025-07-30' => 'Sultan Pahang Hol',
            '2025-08-24' => 'Melaka Governor Birthday',
            '2025-09-14' => 'Sultan Selangor Birthday',
            '2025-10-03' => 'Sabah Governor Birthday',
            '2025-10-13' => 'Sarawak Governor Birthday',

            // ============ 2026 National Holidays ============
            '2026-01-01' => 'New Year\'s Day',
            '2026-02-01' => 'Federal Territory Day / Thaipusam',
            '2026-02-17' => 'Chinese New Year',
            '2026-02-18' => 'Chinese New Year (Day 2)',
            '2026-02-19' => 'Awal Ramadan',
            '2026-03-07' => 'Nuzul Al-Quran',
            '2026-03-21' => 'Hari Raya Aidilfitri',
            '2026-03-22' => 'Hari Raya Aidilfitri (Day 2)',
            '2026-05-01' => 'Labour Day',
            '2026-05-27' => 'Hari Raya Haji',
            '2026-05-28' => 'Hari Raya Haji (Day 2)',
            '2026-05-30' => 'Pesta Kaamatan (Sabah)',
            '2026-05-31' => 'Pesta Kaamatan / Wesak Day',
            '2026-06-01' => 'Gawai Dayak / Agong Birthday',
            '2026-06-02' => 'Gawai Dayak (Sarawak)',
            '2026-06-17' => 'Awal Muharram',
            '2026-07-22' => 'Sarawak Day',
            '2026-08-25' => 'Maulidur Rasul',
            '2026-08-31' => 'National Day',
            '2026-09-16' => 'Malaysia Day',
            '2026-11-08' => 'Deepavali',
            '2026-12-25' => 'Christmas Day',

            // 2026 State Holidays
            '2026-01-14' => 'Hari Hol Negeri Sembilan',
            '2026-01-17' => 'Israk Mikraj',
            '2026-02-20' => 'Hari Pengisytiharan Kemerdekaan (Melaka)',
            '2026-03-04' => 'Hari Ulang Tahun Pertabalan Sultan Terengganu',
            '2026-03-23' => 'Hari Raya (3rd Day) / Sultan Johor Birthday',
            '2026-04-03' => 'Good Friday',
            '2026-04-26' => 'Sultan Terengganu Birthday',
            '2026-05-17' => 'Raja Perlis Birthday',
            '2026-05-22' => 'Hari Hol Pahang',
            '2026-05-26' => 'Hari Arafah',
            '2026-06-21' => 'Sultan Kedah Birthday',
            '2026-07-07' => 'Hari Warisan Dunia Georgetown (Penang)',
            '2026-07-11' => 'Penang Governor Birthday',
            '2026-07-21' => 'Hari Hol Negeri Johor',
            '2026-07-31' => 'Sultan Pahang Birthday',
            '2026-08-24' => 'Melaka Governor Birthday',
            '2026-09-11' => 'Sultan Selangor Birthday',
            '2026-10-03' => 'Sabah Governor Birthday',
            '2026-10-12' => 'Sarawak Governor Birthday',
        ];
    }

    private function loadSchoolHolidays(): array
    {
        $holidays = [];

        // ============ 2025 School Holidays ============
        // Mid-term 1: May 29 - June 9
        $this->addDateRange($holidays, '2025-05-29', '2025-06-09');
        // Mid-term 2: Sept 13 - Sept 21
        $this->addDateRange($holidays, '2025-09-13', '2025-09-21');
        // Year-end: Dec 20, 2025 - Jan 11, 2026
        $this->addDateRange($holidays, '2025-12-20', '2025-12-31');

        // ============ 2026 School Holidays ============
        // Year-end from 2025 continues: Jan 1 - Jan 11
        $this->addDateRange($holidays, '2026-01-01', '2026-01-11');
        // Mid-term 1: March 19 - March 29
        $this->addDateRange($holidays, '2026-03-19', '2026-03-29');
        // Mid-term 2: May 22 - June 6
        $this->addDateRange($holidays, '2026-05-22', '2026-06-06');
        // Mid-term 2 (estimated): mid Sept
        $this->addDateRange($holidays, '2026-09-11', '2026-09-20');
        // Year-end (estimated): mid Dec
        $this->addDateRange($holidays, '2026-12-18', '2026-12-31');

        return $holidays;
    }

    private function addDateRange(array &$holidays, string $start, string $end): void
    {
        $startDate = new \DateTime($start);
        $endDate = new \DateTime($end);
        $endDate->modify('+1 day');

        $interval = new \DateInterval('P1D');
        $period = new \DatePeriod($startDate, $interval, $endDate);

        foreach ($period as $date) {
            $holidays[$date->format('Y-m-d')] = true;
        }
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

                // Check for holiday and school holiday
                $holidayName = null;
                $isHoliday = false;
                $isSchoolHoliday = false;
                if ($showDate && $dateNumber !== null) {
                    $dateKey = sprintf('%04d-%02d-%02d', $year, $month, $dateNumber);
                    if (isset($this->holidays[$dateKey])) {
                        $holidayName = $this->holidays[$dateKey];
                        $isHoliday = true;
                    }
                    if (isset($this->schoolHolidays[$dateKey])) {
                        $isSchoolHoliday = true;
                    }
                }

                // Determine background color with priority handling
                if ($isHoliday && $isSchoolHoliday) {
                    // Public holiday + school holiday = bright yellow
                    $bgColor = self::COLOR_HOLIDAY_SCHOOL_BG;
                } elseif ($isHoliday) {
                    // Public holiday only = pink/red
                    $bgColor = self::COLOR_HOLIDAY_BG;
                } elseif ($isSchoolHoliday && $isWeekend) {
                    // School holiday on weekend = darker yellow-gray
                    $bgColor = self::COLOR_SCHOOL_HOLIDAY_WEEKEND_BG;
                } elseif ($isSchoolHoliday) {
                    // School holiday only = yellow
                    $bgColor = self::COLOR_SCHOOL_HOLIDAY_BG;
                } elseif ($isWeekend) {
                    // Weekend only = gray
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
