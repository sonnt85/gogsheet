package gogsheet

import (
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"sync"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"

	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config, tokFile string) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	// authCode = "4/1AX4XfWg7T0LAGQZKn48HuuLSuGzLywoU8Yju5plaxrtux2uwO1pboVE3xiU"
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

type Gsheet struct {
	mutex                      sync.Mutex
	TokenOauth2_Or_Credentials string
	oauthPath                  string
	spreadsheetId              string
	*sheets.Service
	ctx context.Context
}

func New(oauth2_token_path, credentials_oauth_path, spreadsheetid string) (*Gsheet, error) {
	var err error
	is := &Gsheet{
		TokenOauth2_Or_Credentials: oauth2_token_path,
		oauthPath:                  credentials_oauth_path,
		mutex:                      sync.Mutex{},
		spreadsheetId:              spreadsheetid,
	}

	is.ctx = context.Background()
	if len(oauth2_token_path) == 0 {
		is.Service, err = sheets.NewService(is.ctx, option.WithServiceAccountFile(credentials_oauth_path))
		// is.Service, err = sheets.NewService(is.ctx, option.WithCredentialsFile(credentials_oauth_path))
	} else {
		is.TokenOauth2_Or_Credentials = credentials_oauth_path
		b := make([]byte, 0)
		b, err = ioutil.ReadFile(credentials_oauth_path)
		if err != nil {
			return nil, err
		}
		// If modifying these scopes, delete your previously saved token.json.
		config := new(oauth2.Config)
		config, err = google.ConfigFromJSON(b, sheets.SpreadsheetsScope)
		if err != nil {
			return nil, err
		}
		client := getClient(config, is.TokenOauth2_Or_Credentials)
		is.Service, err = sheets.NewService(is.ctx, option.WithHTTPClient(client))
	}
	if err != nil {
		return nil, err
	}
	return is, nil

}

func (is *Gsheet) UpdateSpreadsheetId(spreadsheetid string) {
	is.spreadsheetId = spreadsheetid
}

func (is *Gsheet) GetValueRange(readRange string, sprids ...string) ([][]string, error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	resp, err := is.Service.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
	if err != nil {
		return nil, err
	}

	if len(resp.Values) == 0 {
		return nil, fmt.Errorf("no data found")
	} else {
		ret := [][]string{}
		for _, row := range resp.Values {
			col := []string{}
			for _, s := range row {
				col = append(col, fmt.Sprint(s))
			}
			ret = append(ret, col)
		}
		return ret, nil
	}
}

func (is *Gsheet) GetValueCell(sheetname, cellAddress string, sprids ...string) (string, error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	if rets, err := is.GetValueRange(fmt.Sprintf("%s!%s:%s", sheetname, cellAddress, cellAddress), spreadsheetId); err == nil {
		if len(rets) != 0 && len(rets[0]) != 0 {
			return rets[0][0], nil
		} else {
			return "", fmt.Errorf("not found")
		}
	} else {
		return "", err
	}
}

func (is *Gsheet) GetValueRanges(readRanges []string, sprids ...string) (map[string][][]string, error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	resp, err := is.Service.Spreadsheets.Values.BatchGet(spreadsheetId).Ranges(readRanges...).Do()
	if err != nil {
		return nil, err
	}

	if len(resp.ValueRanges) == 0 {
		return nil, fmt.Errorf("no data found")
	} else {
		ret := map[string][][]string{}
		for _, valuerange := range resp.ValueRanges {
			retRange := [][]string{}
			for _, row := range valuerange.Values {
				col := []string{}
				for _, s := range row {
					col = append(col, fmt.Sprint(s))
				}
				retRange = append(retRange, col)
			}
			ret[valuerange.Range] = retRange
		}
		return ret, nil
	}
}

func (is *Gsheet) UpdateRanges(rowsArray [][][]interface{}, rangeData []string, sprids ...string) (err error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	// Modify this to your Needs
	batchUpdateValuesRequest := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
	}

	if len(rowsArray) != len(rangeData) {
		return fmt.Errorf("rowsArray and rangeData need same len")
	}
	for i, rows := range rowsArray {
		batchUpdateValuesRequest.Data = append(batchUpdateValuesRequest.Data, &sheets.ValueRange{
			Range:  rangeData[i],
			Values: rows,
		})
	}

	is.mutex.Lock()
	defer is.mutex.Unlock()
	// Do a batch update at once
	_, err = is.Spreadsheets.Values.BatchUpdate(spreadsheetId, batchUpdateValuesRequest).Do()
	return err
}

func (is *Gsheet) UpdateRange(rows [][]interface{}, rangeData string, sprids ...string) (err error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	valueRange := &sheets.ValueRange{
		Values:         rows,
		MajorDimension: "ROWS",
	}
	// Do a batch update at once
	_, err = is.Spreadsheets.Values.Update(spreadsheetId, rangeData, valueRange).ValueInputOption("USER_ENTERED").Do()
	return err
}

func (is *Gsheet) DeleteRange(sheetid int64, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex int64, sprids ...string) (err error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	gridrange := &sheets.DeleteRangeRequest{
		ShiftDimension: "ROWS",
		Range: &sheets.GridRange{
			SheetId: sheetid,
		},
	}
	if startColumnIndex < 0 && endColumnIndex < 0 {
		gridrange.ShiftDimension = "ROWS"
	} else if startRowIndex < 0 && endRowIndex < 0 {
		gridrange.ShiftDimension = "COLUMNS"
	}

	if startColumnIndex >= 0 {
		gridrange.Range.StartColumnIndex = startColumnIndex
	}
	if startRowIndex >= 0 {
		gridrange.Range.StartRowIndex = startRowIndex
	}
	if endRowIndex >= 0 {
		gridrange.Range.EndRowIndex = endRowIndex
	}
	if endColumnIndex >= 0 {
		gridrange.Range.EndColumnIndex = endColumnIndex
	}
	rq := &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: true,
		Requests:                     []*sheets.Request{&sheets.Request{DeleteRange: gridrange}},
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	_, err = is.Spreadsheets.BatchUpdate(spreadsheetId, rq).Do()
	return err
}

func (is *Gsheet) ClearRange(rangeA1 string, sprids ...string) (err error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	_, err = is.Spreadsheets.Values.Clear(spreadsheetId, rangeA1, new(sheets.ClearValuesRequest)).Do()
	return err
}

func (is *Gsheet) ClearRanges(sheetid int64, rangesA1 []string, sprids ...string) (err error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	_, err = is.Spreadsheets.Values.BatchClear(spreadsheetId, &sheets.BatchClearValuesRequest{Ranges: rangesA1}).Do()
	return err
}

func (is *Gsheet) AppendRows(rows [][]interface{}, rangeData string, sprids ...string) (err error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	// Modify this to your Needs
	valueRange := &sheets.ValueRange{
		Values: rows,
		// MajorDimension: "ROWS",
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	// Do a value append at once
	_, err = is.Spreadsheets.Values.Append(spreadsheetId, rangeData, valueRange).ValueInputOption("USER_ENTERED").Do()
	return err
}

func (is *Gsheet) ListSheets(sprids ...string) (map[string]int64, error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	resp, err := is.Spreadsheets.Get(spreadsheetId).Do()
	if err != nil {
		log.Fatal(err)
	}
	ret := map[string]int64{}
	for _, v := range resp.Sheets {
		ret[v.Properties.Title] = v.Properties.SheetId
	}
	return ret, nil
}

func (is *Gsheet) GetSheetIdFromNAme(sheetName string, sprids ...string) (int64, error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	if mapsheets, err := is.ListSheets(spreadsheetId); err == nil {
		if sheetidInt, ok := mapsheets[sheetName]; ok {
			return sheetidInt, nil
		} else {
			return 0, fmt.Errorf("not found")
		}
	} else {
		return 0, err
	}
}

func (is *Gsheet) CreaateSheet(nameSheet string, sprids ...string) (int64, error) {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	rq := &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: true,
		Requests:                     []*sheets.Request{&sheets.Request{AddSheet: &sheets.AddSheetRequest{Properties: &sheets.SheetProperties{Title: nameSheet}}}},
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	respone, err := is.Spreadsheets.BatchUpdate(spreadsheetId, rq).Do()
	if err != nil {
		return 0, err
	}
	for _, v := range respone.UpdatedSpreadsheet.Sheets {
		if v.Properties.Title == nameSheet {
			return v.Properties.SheetId, nil
		}
	}
	return 0, fmt.Errorf("can not found sheet after creat")
}

func (is *Gsheet) DeleteSheetId(sheetid int64, sprids ...string) error {
	spreadsheetId := is.spreadsheetId
	if len(sprids) != 0 {
		spreadsheetId = sprids[0]
	}
	rq := &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: false,
		Requests:                     []*sheets.Request{&sheets.Request{DeleteSheet: &sheets.DeleteSheetRequest{SheetId: sheetid}}},
	}
	is.mutex.Lock()
	defer is.mutex.Unlock()
	_, err := is.Spreadsheets.BatchUpdate(spreadsheetId, rq).Do()
	return err
}

func (is *Gsheet) DeleteSheetFromName(sheetid string, sprids ...string) error {
	lsheets, err := is.ListSheets(sprids...)
	if err != nil {
		return err
	}

	id, ok := lsheets[sheetid]
	if ok {
		return is.DeleteSheetId(id, sprids...)
	} else {
		return fmt.Errorf("can not find sheetid %s", sheetid)
	}
}
