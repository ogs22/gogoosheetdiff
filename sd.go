package main

import (
	"encoding/json"
	"fmt"
	"github.com/sergi/go-diff/diffmatchpatch"
	"google.golang.org/api/sheets/v4"
	"log"
	"os"
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

	//the old spreadsheet id
	//gs_old := "1qRyEu-NU1lAJWuQ1V4Q9PBF13DnqTZXIQsYkphedjqg"
	//gs_new := "1gl6DFGuBM66bsb_bRno6PneITIJr7M9pEpdCtdmyCbI"

	gs_old := "1WVDcBOCVJOGMqMiNBxUoWJVAhUkQhHQ4TJSoRP1V9DA"
	gs_new := "1Cfwi9wlYSAIp69MmRqga1EAgAeCDh-T-Gb-LbIaf0Ss"

	fmt.Printf("\nComparing sheet %s and %s", gs_old, gs_new)

	//get *sheets.Service from quickstart
	service := qsmain()
	// get the props for the old sheets
	sother_old, err := service.Spreadsheets.Get(gs_old).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}
	tmp_old, err := sother_old.MarshalJSON()
	err = json.Unmarshal(tmp_old, &gs_old_props)
	if err != nil {
		log.Fatalf("Unable to Unmarshall from sheet: %v", err)
	}
	sother_new, err := service.Spreadsheets.Get(gs_new).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}
	tmp_new, err := sother_new.MarshalJSON()
	err = json.Unmarshal(tmp_new, &gs_new_props)
	if err != nil {
		log.Fatalf("Unable to Unmarshall from sheet: %v", err)
	}
	_ = cmpSheetTitles(gs_old_props, gs_new_props)

	old_content := getSheetContents(gs_old_props, service)
	new_content := getSheetContents(gs_new_props, service)

	dmp := diffmatchpatch.New()
	for k, _ := range old_content {
		fmt.Printf("\nINFO: Checking %s\n", k)
		diffs := dmp.DiffMain(old_content[k], new_content[k], true)
		//fmt.Println(dmp.DiffPrettyHtml(diffs))

		fmt.Println(dmp.DiffPrettyText(diffs))
		//fmt.Println(dmp.DiffToDelta(diffs))
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
