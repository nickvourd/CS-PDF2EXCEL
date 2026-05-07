# CS-PDF2EXCEL

<p align="center">
  <img src="/Pictures/banner.svg" alt="CS-PDF2EXCEL banner" width="100%"/>
</p>

Convert a [Cobalt Strike](https://www.cobaltstrike.com/) **Sessions Report** PDF into a clean, multi-sheet Excel workbook for filtering, pivoting, and reporting.

## Features

- Skips the first page automatically
- Strips per-page headers (`Sessions Report`) and footers (`Page. N`)
- Ignores all images and icons (text-only extraction)
- Reconstructs multi-line table cells (e.g. wrapped file paths)
- Writes four normalized sheets joinable via `session_id`:
  - `Sessions` — hostname, user, process, pid, opened
  - `CommunicationPaths` — host, port, protocol
  - `FileHashes` — date, hash, name
  - `Activities` — date, activity

## Requirements

- Python 3.9+
- [pdfplumber](https://pypi.org/project/pdfplumber/)
- [openpyxl](https://pypi.org/project/openpyxl/)

## Installation

```bash
git clone https://github.com/nickvourd/CS-PDF2EXCEL.git
cd CS-PDF2EXCEL
pip install -r requirements.txt
```

## Usage

```bash
python pdf_to_excel.py <input.pdf> [output.xlsx]
```

If the output path is omitted, the workbook is written next to the input PDF using the same base name (e.g. `sessionsreport.pdf` -> `sessionsreport.xlsx`).

### Examples

```bash
# Default output (sessionsreport.xlsx alongside the PDF)
python pdf_to_excel.py sessionsreport.pdf

# Explicit output path
python pdf_to_excel.py sessionsreport.pdf out/report.xlsx
```

### Sample output

```
 ██████╗███████╗      ██████╗ ██████╗ ███████╗██████╗ ███████╗██╗  ██╗ ██████╗███████╗██╗
██╔════╝██╔════╝      ██╔══██╗██╔══██╗██╔════╝╚════██╗██╔════╝╚██╗██╔╝██╔════╝██╔════╝██║
██║     ███████╗█████╗██████╔╝██║  ██║█████╗   █████╔╝█████╗   ╚███╔╝ ██║     █████╗  ██║
██║     ╚════██║╚════╝██╔═══╝ ██║  ██║██╔══╝  ██╔═══╝ ██╔══╝   ██╔██╗ ██║     ██╔══╝  ██║
╚██████╗███████║      ██║     ██████╔╝██║     ███████╗███████╗██╔╝ ██╗╚██████╗███████╗███████╗
 ╚═════╝╚══════╝      ╚═╝     ╚═════╝ ╚═╝     ╚══════╝╚══════╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝
                          Created with <3 by @nickvourd

Wrote sessionsreport.xlsx
  Sessions:            16
  Communication paths: 16
  File hashes:         3
  Activities:          4604
```

## Output schema

| Sheet                | Columns                                                |
| -------------------- | ------------------------------------------------------ |
| `Sessions`           | session_id, hostname, user, process, pid, opened       |
| `CommunicationPaths` | session_id, hostname, host, port, protocol             |
| `FileHashes`         | session_id, hostname, date, hash, name                 |
| `Activities`         | session_id, hostname, date, activity                   |

`session_id` is assigned in document order and is consistent across sheets, so you can pivot or join on it directly in Excel / Power Query / pandas.

## License

MIT

## Author

Created with <3 by [@nickvourd](https://github.com/nickvourd)
