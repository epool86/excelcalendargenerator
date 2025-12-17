<?php

require_once __DIR__ . '/vendor/autoload.php';

use App\CalendarGenerator;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Handle form submission
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['year'])) {
    $year = (int) $_POST['year'];

    // Validate year
    if ($year < 2020 || $year > 2100) {
        die('Invalid year selected');
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

// Default selected year
$selectedYear = 2026;
$currentYear = (int) date('Y');
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Calendar Generator</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1B4F72 0%, #5B2C6F 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }

        .container {
            background: white;
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            text-align: center;
            max-width: 450px;
            width: 100%;
        }

        h1 {
            color: #1B4F72;
            margin-bottom: 10px;
            font-size: 28px;
        }

        .subtitle {
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }

        .form-group {
            margin-bottom: 25px;
        }

        label {
            display: block;
            color: #333;
            font-weight: 600;
            margin-bottom: 10px;
            font-size: 16px;
        }

        select {
            width: 100%;
            padding: 15px 20px;
            font-size: 18px;
            border: 2px solid #ddd;
            border-radius: 8px;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s ease;
        }

        select:focus {
            outline: none;
            border-color: #5B2C6F;
        }

        button {
            width: 100%;
            padding: 15px 30px;
            font-size: 18px;
            font-weight: 600;
            color: white;
            background: linear-gradient(135deg, #5B2C6F 0%, #1B4F72 100%);
            border: none;
            border-radius: 8px;
            cursor: pointer;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(91, 44, 111, 0.4);
        }

        button:active {
            transform: translateY(0);
        }

        .features {
            margin-top: 30px;
            padding-top: 25px;
            border-top: 1px solid #eee;
            text-align: left;
        }

        .features h3 {
            color: #1B4F72;
            font-size: 14px;
            margin-bottom: 12px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .features ul {
            list-style: none;
            color: #666;
            font-size: 13px;
        }

        .features li {
            padding: 5px 0;
            padding-left: 20px;
            position: relative;
        }

        .features li::before {
            content: '✓';
            position: absolute;
            left: 0;
            color: #27ae60;
            font-weight: bold;
        }

        .flag {
            font-size: 24px;
            margin-bottom: 15px;
        }

        .footer {
            margin-top: 25px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            font-size: 12px;
            color: #888;
        }

        .footer a {
            color: #5B2C6F;
            text-decoration: none;
        }

        .footer a:hover {
            text-decoration: underline;
        }

        .footer p {
            margin: 5px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="flag">🇲🇾</div>
        <h1>Excel Calendar Generator</h1>
        <p class="subtitle">Generate yearly calendar with Malaysian public holidays</p>

        <form method="POST" action="">
            <div class="form-group">
                <label for="year">Select Year</label>
                <select name="year" id="year">
                    <?php for ($y = $currentYear; $y <= $currentYear + 5; $y++): ?>
                        <option value="<?= $y ?>" <?= $y === $selectedYear ? 'selected' : '' ?>>
                            <?= $y ?>
                        </option>
                    <?php endfor; ?>
                </select>
            </div>

            <button type="submit">
                Generate & Download Excel
            </button>
        </form>

        <div class="features">
            <h3>Features</h3>
            <ul>
                <li>12 monthly sheets in one file</li>
                <li>Malaysian public holidays highlighted</li>
                <li>Space to write daily activities</li>
                <li>Weekend columns marked in grey</li>
                <li>Professional formatting & styling</li>
            </ul>
        </div>

        <div class="footer">
            <p>Created by <a href="https://fb.com/asbahri" target="_blank">AHMAD SAIFUL BAHRI</a></p>
            <p><a href="https://github.com/epool86/excelcalendargenerator" target="_blank">View Source Code on GitHub</a></p>
        </div>
    </div>
</body>
</html>
