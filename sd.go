package main

import (
	"encoding/json"
	"fmt"
	"github.com/fatih/color"
	"github.com/sergi/go-diff/diffmatchpatch"
	"google.golang.org/api/sheets/v4"
	"io/ioutil"
	"log"
	"net/http"
	"os"

	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"regexp"
)

type MySheet struct {
	SpreadsheetID string `json:"spreadsheetId"`
	Properties    struct {
		Title         string `json:"title"`
		Locale        string `json:"locale"`
		AutoRecalc    string `json:"autoRecalc"`
		TimeZone      string `json:"timeZone"`
		DefaultFormat struct {
			BackgroundColor struct {
				Red   int `json:"red"`
				Green int `json:"green"`
				Blue  int `json:"blue"`
			} `json:"backgroundColor"`
			Padding struct {
				Top    int `json:"top"`
				Right  int `json:"right"`
				Bottom int `json:"bottom"`
				Left   int `json:"left"`
			} `json:"padding"`
			VerticalAlignment string `json:"verticalAlignment"`
			WrapStrategy      string `json:"wrapStrategy"`
			TextFormat        struct {
				ForegroundColor struct {
				} `json:"foregroundColor"`
				FontFamily    string `json:"fontFamily"`
				FontSize      int    `json:"fontSize"`
				Bold          bool   `json:"bold"`
				Italic        bool   `json:"italic"`
				Strikethrough bool   `json:"strikethrough"`
				Underline     bool   `json:"underline"`
			} `json:"textFormat"`
		} `json:"defaultFormat"`
	} `json:"properties"`
	Sheets []struct {
		Properties struct {
			SheetID        int    `json:"sheetId"`
			Title          string `json:"title"`
			Index          int    `json:"index"`
			SheetType      string `json:"sheetType"`
			GridProperties struct {
				RowCount    int `json:"rowCount"`
				ColumnCount int `json:"columnCount"`
			} `json:"gridProperties"`
		} `json:"properties"`
	} `json:"sheets"`
	SpreadsheetURL string `json:"spreadsheetUrl"`
}

func main() {
	var gs_old_props, gs_new_props MySheet

	info := color.New(color.FgCyan).Add(color.Underline)

	if len(os.Args) < 2 {
		log.Printf("Usage %s SpreadSheetID_old SpreadSheetID_new", os.Args[0])
		log.Fatalf("Incorrect command line args")
	}

	gs_old := os.Args[1]
	gs_new := os.Args[2]

	var validID = regexp.MustCompile(`[a-zA-Z0-9-_]+`)

	if validID.MatchString(gs_old) && validID.MatchString(gs_new) {
		// ids are valid
		info.Printf("\nComparing sheet %s and %s\n", gs_old, gs_new)
	} else {
		log.Fatalf("Input IDs not valid \n%s \n%s", gs_old ,gs_new)
	}



	//get *sheets.Service from quickstart
	service := qsmain()
	// get the props for the old sheets
	sother_old, err := service.Spreadsheets.Get(gs_old).Do()
	if err != nil {
		log.Fatalf("\nUnable to retrieve data from spreadsheet\nID:%s: %v",gs_old, err)
	}
	tmp_old, err := sother_old.MarshalJSON()
	err = json.Unmarshal(tmp_old, &gs_old_props)
	if err != nil {
		log.Fatalf("\nUnable to Unmarshall from spreadsheet\nID:%s: %v",gs_old, err)
	}
	sother_new, err := service.Spreadsheets.Get(gs_new).Do()
	if err != nil {
		log.Fatalf("\nUnable to retrieve data from spreadsheet\nID:%s: %v",gs_new, err)
	}
	tmp_new, err := sother_new.MarshalJSON()
	err = json.Unmarshal(tmp_new, &gs_new_props)
	if err != nil {
		log.Fatalf("\nUnable to Unmarshall from spreadsheet\nID:%s: %v",gs_new, err)
	}
	info.Println("\nChecking Spreadsheets have the same number of sheets and sheet names")
	_ = cmpSheetTitles(gs_old_props, gs_new_props)

	old_content := getSheetContents(gs_old_props, service)
	new_content := getSheetContents(gs_new_props, service)

	dmp := diffmatchpatch.New()
	for k, _ := range old_content {
		info.Printf("\nINFO: Checking %s\n", k)
		diffs := dmp.DiffMain(old_content[k], new_content[k], true)
		fmt.Println(dmp.DiffPrettyText(diffs))
	}
}

func getSheetContents(props MySheet, service *sheets.Service) map[string]string {
	contents := make(map[string]string)
	valueRenderOption := "FORMULA"
	for _, row := range props.Sheets {
		this_title := row.Properties.Title
		resp, err := service.Spreadsheets.Values.Get(props.SpreadsheetID, this_title).ValueRenderOption(valueRenderOption).Do()
		if err != nil {
			log.Fatalf("Unable to retrieve data from sheet: %v", err)
		}
		for _, brow := range resp.Values {
			for _, col := range brow {
				contents[this_title] += "'" + col.(string) + "',"
			}
			contents[this_title] += "\n"
		}
	}
	return contents
}

//compares 2 spreadsheet's sheet titles and exits if they are not the same
// and in the same order
func cmpSheetTitles(old_props MySheet, new_props MySheet) bool {
	var old_titles, new_titles []string

	for _, row := range old_props.Sheets {
		old_titles = append(old_titles, row.Properties.Title)
	}
	for i, row := range new_props.Sheets {
		new_titles = append(new_titles, row.Properties.Title)
		if row.Properties.Title != old_titles[i] {
			fmt.Printf("\n%s != %s", row.Properties.Title, old_titles[i])
			os.Exit(1)
			return false
		}
		fmt.Printf("\n%s Exists in both Spreadsheets", row.Properties.Title)
	}
	if len(old_titles) != len(new_titles) {
		fmt.Println("\nSpreadsheets have different number of Sheets")
		fmt.Printf("\nThe old Spreadsheet has %d and the new %d", len(old_titles), len(new_titles))
		os.Exit(1)
		return false
	}

	fmt.Printf("\nINFO:%s", "Sheets have the same titles")
	return true
}

// From Google api quickstart https://github.com/gsuitedevs/go-samples/blob/master/sheets/quickstart/quickstart.go
// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	tokFile := "token.json"
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
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(oauth2.NoContext, authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	defer f.Close()
	if err != nil {
		return nil, err
	}
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	defer f.Close()
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	json.NewEncoder(f).Encode(token)
}

func qsmain()(*sheets.Service) {
	b, err := ioutil.ReadFile("credentials.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets.readonly")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)

	srv, err := sheets.New(client)

	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}
	return srv
}

// [END sheets_quickstart]
