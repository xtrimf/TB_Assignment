package main

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"regexp"
	"strings"
	"time"

	"golang.org/x/oauth2/google"
	"gopkg.in/Iwark/spreadsheet.v2"
)

// prepare RegEx templates
var macRegExp = regexp.MustCompile(`(?:[a-z,A-Z,0-9]{4}.[a-z,A-Z,0-9]{4}?.[a-z,A-Z,0-9]{4}|([a-z,A-Z,0-9]{2}[:.]){5}[a-z,A-Z,0-9]{2})`) //mac address validation in form of "xxxx.xxxx.xxxx" OR "xx:xx:xx:xx:xx:xx" - no checking for Hexadecimal chars
var ipRegExp = regexp.MustCompile(`\d{1,3}\.\d{1,3}\.\d{1,3}\.[0-9Z]{1,3}`)                                                             // ip address validation - exception for "Z" chars in 4th oct
var serialRegExp = regexp.MustCompile(`[A-Z,0-9]{7,8}`)                                                                                 //serial number validation
var template1 = regexp.MustCompile(`SW - .{4}\..{4}\..{4}`)                                                                             // pattern #1 - SW with negative outlook of tags
var template2 = regexp.MustCompile(`\n\d{1,3}\.\d{1,3}\.\d{1,3}\.[0-9Z]{1,3}`)                                                          // pattern #2 - multiline
var template3 = regexp.MustCompile(`PDU [A-B] - (BLUE|RED)`)                                                                            // pattern #3 - PDU with details
var template4 = regexp.MustCompile(`SW.+?TAG:`)                                                                                         // pattern #4 - SW with tags

// Row - final data structure
type Row struct {
	Name     string `json:"Name"`
	IP       string `json:"ipmi"`
	MAC      string `json:"MAC"`
	Serial   string `json:"serial"`
	Location string `json:"Location"`
}

// Rows - array of row
var Rows = []Row{}

// Cell - cells of raw data from the gsheet
type Cell struct {
	row    uint
	col    uint
	value  string
	header uint
	h      string
}

func main() {
	start := time.Now().UTC()

	service := authenticate()
	spreadsheet, err := service.FetchSpreadsheet("13ITAyM-8pq9hay_Y2NNOqS_ps2rSlWR080yBEEviXmI")
	checkError(err)

	//start looping over sheets
	headers := []string{"Electricity", "Name", "IPMI", "MAC", "Serial"}

	println("Main: Processing...")

	for _, sheet := range spreadsheet.Sheets {
		getSheetData(sheet, headers[0])
	}

	fmt.Printf("Main: Completed. Results (%v lines) are stored in 'Output.json'.", len(Rows))
	wrightResults()

	fmt.Printf("\nProcessing time (sec): %.3f", time.Since(start).Seconds())
}

func template(data string) int {
	if template1.MatchString(data) && !template4.MatchString(data) {
		return 1
	}
	if template2.MatchString(data) {
		return 2
	}
	if template3.MatchString(data) {
		return 3
	}
	if template4.MatchString(data) {
		return 4
	}
	return 0 // no pattern - single cell distribution
}

func matchCell(data string) int {
	if ipRegExp.MatchString(data) {
		return 1
	}
	if macRegExp.MatchString(data) {
		return 2
	}
	return 0
}

func getSheetData(sheet spreadsheet.Sheet, matchHeader string) {
	var cellList = []Cell{}
	var h1, h2, headerRow, lastCol uint
	for _, row := range sheet.Rows {
		for _, cell := range row {
			if cell.Value == matchHeader {
				if h1 == 0 {
					h1 = cell.Column // mark first "half"
					h2 = h1 + 5      // mark second "half"
					lastCol = h2 + 4
					headerRow = cell.Row
				}
			}
			if h1 > 0 && headerRow < cell.Row && h1 <= cell.Column && lastCol >= cell.Column && cell.Value != "empty" && cell.Value != "" {
				newCell := Cell{
					row:    cell.Row,
					col:    cell.Column,
					value:  cell.Value,
					header: cell.Column,
				}
				if cell.Column < h2 {
					newCell.h = "h1"
				} else {
					newCell.h = "h2"
				}

				cellList = append(cellList, newCell)
			}
		}
	}
	buidRows(sheet.Properties.Title, cellList)
}

func buidRows(sheetName string, cellList []Cell) {
	for i := 0; i < len(cellList); i++ {
		switch template := template(cellList[i].value); template {
		case 1:
			splitRow := strings.Split(cellList[i].value, " - ")
			row := Row{
				Name:     splitRow[0],
				IP:       splitRow[2],
				MAC:      splitRow[1],
				Serial:   splitRow[3],
				Location: sheetName,
			}
			Rows = append(Rows, row)
		case 2:
			splitRow := strings.Split(cellList[i].value, "\n")
			row := Row{
				Name:     splitRow[0],
				Serial:   cellList[i+1].value,
				Location: sheetName,
			}
			assignIPMAC(cellList[i].value, &row)
			Rows = append(Rows, row)
			i++ //advance one cell
		case 3:
			row := Row{
				Name:     template3.FindString(cellList[i].value),
				Serial:   "",
				Location: sheetName,
			}
			assignIPMAC(cellList[i].value, &row)
			Rows = append(Rows, row)
		case 4:
			row := Row{
				Name:     strings.TrimSpace(strings.Split(cellList[i].value, "SW")[0] + "SW"),
				Serial:   serialRegExp.FindString(cellList[i].value),
				Location: sheetName,
			}
			assignIPMAC(cellList[i].value, &row)
			Rows = append(Rows, row)
		default:
			row := Row{
				Name:     strings.TrimSpace(cellList[i].value),
				Location: sheetName,
			}
			currentH := cellList[i].h
			currentRow := cellList[i].row
			// match the rest of the cells after "name"
			if i < len(cellList)-1 && cellList[i+1].h == currentH && cellList[i+1].row == currentRow {
				i++
				for {
					if i < len(cellList)-1 && cellList[i].h == currentH && cellList[i].row == currentRow {
						if ipRegExp.MatchString(cellList[i].value) {
							row.IP = ipRegExp.FindString(cellList[i].value)
						} else if macRegExp.MatchString(cellList[i].value) {
							row.MAC = macRegExp.FindString(cellList[i].value)
						} else {
							row.Serial = cellList[i].value
						}
						if cellList[i+1].h == currentH && cellList[i+1].row == currentRow {
							i++
						} else {
							break
						}
					} else {
						break
					}
				}
			}
			Rows = append(Rows, row)
			//fmt.Printf("%+v\n", row)
		}
	}

}

// reusable function to assign ip & MAC
func assignIPMAC(value string, row *Row) {
	row.IP = ipRegExp.FindString(value)
	row.MAC = macRegExp.FindString(value)
}

// function to authenticate on Google
func authenticate() *spreadsheet.Service {
	data, err := ioutil.ReadFile("client_secret.json")
	checkError(err)
	conf, err := google.JWTConfigFromJSON(data, spreadsheet.Scope)
	checkError(err)
	client := conf.Client(context.TODO())
	service := spreadsheet.NewServiceWithClient(client)
	return service
}

// function to write results to a file
func wrightResults() {
	rowsJSON, _ := json.Marshal(Rows)
	err := ioutil.WriteFile("output.json", bytes.ReplaceAll(rowsJSON, []byte("},{"), []byte("},\n{")), 0644)
	checkError(err)
}

func checkError(err error) {
	if err != nil {
		panic(err.Error())
	}
}
