# Excelibur

## CAUTION: THIS PROJECT IS STILL A WORK IN PROGRESS, AND MIGHT NOT WORK AS EXPECTED (OR AT ALL!)

**Pure Go XLS to XLSX converter with zero dependencies**

Excelibur is a lightweight Go package that converts XLS (Excel 97-2003) files to XLSX format by directly parsing the binary file structure, with no external dependencies or requirements.

[![Go Report Card](https://goreportcard.com/badge/github.com/NickTacke/excelibur)](https://goreportcard.com/report/github.com/NickTacke/excelibur)
[![GoDoc](https://godoc.org/github.com/NickTacke/excelibur?status.svg)](https://godoc.org/github.com/NickTacke/excelibur)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **Zero External Dependencies**: No need for LibreOffice, Python, or other external tools
- **Pure Go Implementation**: Parse XLS binary format natively in Go
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Stream Processing**: Low memory usage even with large files
- **Comprehensive Format Support**:
  - Multiple worksheets
  - Text, numbers, dates, and boolean values
  - Basic formulas (preserved as text)
  - Shared strings optimization

## Installation

### Library

```bash
go get github.com/NickTacke/excelibur
```

### Command-line Tool

```bash
go install github.com/NickTacke/excelibur/cmd/excelibur@latest
```

## Usage

### As a Library

```go
package main

import (
    "fmt"
    "github.com/NickTacke/excelibur"
)

func main() {
    // Create a new converter
    converter := excelibur.NewConverter()
    
    // Convert a single file
    err := converter.ConvertFile("input.xls", "output.xlsx")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
    }
    
    // Convert all files in a directory
    files, err := converter.ConvertDirectory("./docs", "./converted")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
    }
    
    fmt.Printf("Successfully converted %d files\n", len(files))
}
```

### Command-line Tool

```bash
# Convert a single file
excelibur -input file.xls -output converted.xlsx

# Convert all XLS files in a directory
excelibur -input ./docs -output ./converted

# Show conversion details
excelibur -input file.xls -verbose
```

## Technical Details

Excelibur works by directly parsing the OLE2 Compound File Binary Format (the container format used by XLS files) and the BIFF8 (Binary Interchange File Format) records within.

### Supported Binary Structures

- **OLE2 Compound File Format**: The container format used by Microsoft Office
- **BIFF8 Records**: The binary format used by Excel 97-2003
- **Shared String Table**: For efficient text storage
- **Cell Records**: Various cell types (RK, NUMBER, BOOLERR, etc.)
- **Worksheet Structures**: Dimensions, rows, cells, etc.

### Binary Parsing Process

1. **OLE2 Container**: Parse the compound file structure to locate the Workbook stream
2. **Workbook Stream**: Extract global information and sheet data locations
3. **Worksheet Streams**: Parse individual worksheets and their cells
4. **XLSX Generation**: Convert the parsed data to XLSX using the Go standard library

## Performance

Excelibur is designed to be efficient and handle large XLS files with minimal memory usage:

- **Memory Efficient**: Processes files in a streaming fashion
- **Fast**: Typically converts files in milliseconds to seconds, depending on size
- **Low CPU Usage**: Efficient binary parsing with minimal overhead

## Limitations

- Limited support for advanced Excel features (macros, pivot tables, charts)
- Some complex formatting might not be preserved
- Comments, drawings, and other objects have limited support
- Password-protected files are not supported

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- The [OpenOffice Documentation of the Microsoft Excel File Format](http://www.openoffice.org/sc/excelfileformat.pdf)
- Microsoft's [MS-XLS](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls) specification
- Microsoft's [MS-CFB](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb) specification
