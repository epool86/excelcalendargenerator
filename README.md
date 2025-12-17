# Excel Calendar Generator

A simple PHP web application that generates downloadable Excel calendar files with Malaysian public holidays.

## Live Demo

**[https://calendar.azbahri.link](https://calendar.azbahri.link)**

## Features

- 12 monthly sheets in one Excel file
- Malaysian public holidays (2025-2026) auto-populated
- Horse-race style calendar layout
- Space to write daily activities
- Weekend columns highlighted in grey
- Holiday cells styled with pink background and red text
- Professional formatting with borders

## Requirements

- PHP 8.1+
- Composer
- PHP extensions: `zip`, `gd` or `imagick`

## Installation

```bash
git clone git@github.com:epool86/excelcalendargenerator.git
cd excelcalendargenerator
composer install
```

## Usage

### Development

```bash
php -S localhost:8000
```

Then open `http://localhost:8000` in your browser.

### Production

Point your web server (Apache/Nginx) to the project root directory.

## Creator

**AHMAD SAIFUL BAHRI**

- Facebook: [fb.com/asbahri](https://fb.com/asbahri)
- GitHub: [github.com/epool86](https://github.com/epool86)

## License

MIT License
