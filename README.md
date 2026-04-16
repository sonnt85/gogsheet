# gogsheet

A Google Sheets API client for reading, writing, and managing spreadsheet data in Go.

## Installation

```bash
go get github.com/sonnt85/gogsheet
```

## Features

- Authenticate via OAuth2 token or service account credentials file
- Read cell values and ranges (single and batch)
- Update ranges with batch or single operations
- Append rows to a sheet
- Clear single or multiple ranges
- Delete ranges (shift rows or columns)
- List, create, and delete sheets within a spreadsheet
- Thread-safe operations via mutex

## Usage

```go
package main

import (
    "fmt"
    "log"

    "github.com/sonnt85/gogsheet"
)

func main() {
    // Connect using a service account (empty oauth2 token path)
    gs, err := gogsheet.New("", "service-account.json", "SPREADSHEET_ID")
    if err != nil {
        log.Fatal(err)
    }

    // Read a range
    rows, err := gs.GetValueRange("Sheet1!A1:C10")
    if err != nil {
        log.Fatal(err)
    }
    for _, row := range rows {
        fmt.Println(row)
    }

    // Read a single cell
    val, _ := gs.GetValueCell("Sheet1", "B2")
    fmt.Println(val)

    // Append rows
    gs.AppendRows([][]interface{}{{"col1", "col2"}}, "Sheet1")

    // Create and delete sheets
    sheetID, _ := gs.CreateSheet("NewSheet")
    fmt.Println("Created sheet ID:", sheetID)
    gs.DeleteSheetId(sheetID)
}
```

## API

- `New(oauth2TokenPath, credentialsPath, spreadsheetID string) (*Gsheet, error)` — creates a new client
- `(*Gsheet).UpdateSpreadsheetId(id string)` — changes the active spreadsheet ID
- `(*Gsheet).GetValueRange(readRange string, sprids ...string) ([][]string, error)` — reads a range
- `(*Gsheet).GetValueCell(sheetname, cellAddress string, sprids ...string) (string, error)` — reads a single cell
- `(*Gsheet).GetValueRanges(readRanges []string, sprids ...string) (map[string][][]string, error)` — batch reads multiple ranges
- `(*Gsheet).UpdateRange(rows [][]interface{}, rangeData string, sprids ...string) error` — updates a range
- `(*Gsheet).UpdateRanges(rowsArray [][][]interface{}, rangeData []string, sprids ...string) error` — batch updates
- `(*Gsheet).AppendRows(rows [][]interface{}, rangeData string, sprids ...string) error` — appends rows
- `(*Gsheet).ClearRange(rangeA1 string, sprids ...string) error` — clears a range
- `(*Gsheet).ClearRanges(sheetid int64, rangesA1 []string, sprids ...string) error` — batch clears
- `(*Gsheet).DeleteRange(sheetid, startRow, startCol, endRow, endCol int64, sprids ...string) error` — deletes a range
- `(*Gsheet).ListSheets(sprids ...string) (map[string]int64, error)` — lists all sheets
- `(*Gsheet).GetSheetIdFromName(sheetName string, sprids ...string) (int64, error)` — looks up a sheet ID by name
- `(*Gsheet).CreateSheet(nameSheet string, sprids ...string) (int64, error)` — creates a new sheet
- `(*Gsheet).DeleteSheetId(sheetid int64, sprids ...string) error` — deletes a sheet by ID
- `(*Gsheet).DeleteSheetFromName(sheetName string, sprids ...string) error` — deletes a sheet by name

## Author

**sonnt85** — [thanhson.rf@gmail.com](mailto:thanhson.rf@gmail.com)

## License

MIT License - see [LICENSE](LICENSE) for details.
