package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

const (
	submitterNameCell       = "A44"
	dateCell                = "A45"
	monthBeginCol           = 'D'
	dateFormat              = "02-01-2006"
	generatedFileDateFormat = "02-Jan-2006"
	monthFormat             = "Jan"
	dayOfMonthFormat        = "02/01"
	daysEarned              = "2.5"
	outputDir               = "gen/"
)

const (
	months                = 12
	monthRow              = 7
	initialsRow           = 42
	supervisorInitialsRow = 43
	dayOfMonthRow         = 44
	startingBalanceRow    = 46
	daysEarnedRow         = 47
	leaveRow              = 48
	newBalanceRow         = 49
)

var sheet string

var template = flag.String("template", "timesheet.xlsx", "Timesheet template file")
var submitter = flag.String("name", "C. Gurure", "Submitter Name")
var initials = flag.String("initials", "CG", "Submitter Initials")
var supervisorInitials = flag.String("sinitials", "TD", "Supervisor Initials")
var sheetIndex = flag.Int("sheet", 0, "0-based Sheet index")
var month = flag.Int("month", int(time.Now().Month()), "1-based Month number")
var year = flag.Int("year", time.Now().Year(), "Year")
var leave = flag.Int("leave", 0, "Number of leave days")

func lastDayOfMonth(fromDate time.Time) time.Time {
	return time.Date(fromDate.Year(), fromDate.Month()+1, 0, 0, 0, 0, 0, time.UTC)
}

func statFile(f string) error {
	_, err := os.Stat(f)
	return err
}

func openSheet(fname string) (*excelize.File, error) {
	if err := statFile(fname); err != nil {
		return nil, fmt.Errorf("input template unreadable: %v", err)
	}
	f, err := excelize.OpenFile(fname)
	if err != nil {
		return nil, fmt.Errorf("failed to read input template: %v\n", err)
	}
	return f, nil
}

func genTemplateName(submissionDate time.Time) string {
	return fmt.Sprintf("%stimesheet-%s.xlsx", outputDir, submissionDate.Format(generatedFileDateFormat))
}

func saveSheet(f *excelize.File, submissionDate time.Time) error {
	updatedFile := genTemplateName(submissionDate)
	return f.SaveAs(updatedFile)
}

func monthCol(offset int) rune {
	colOffset := int(monthBeginCol) + offset
	return rune(colOffset)
}

func currentTime(m int) time.Time {
	now := time.Now()
	return time.Date(*year, time.Month(m), now.Day(), now.Hour(),
		now.Minute(), now.Second(), now.Nanosecond(), time.UTC)
}

func clearInitials(f *excelize.File, col string) {
	f.SetCellValue(sheet, cell(col, initialsRow), "")
	f.SetCellValue(sheet, cell(col, supervisorInitialsRow), "")
	f.SetCellValue(sheet, cell(col, dayOfMonthRow), "")
}

func updateInitials(f *excelize.File, col string) {

	f.SetCellValue(sheet, cell(col, initialsRow), *initials)
	f.SetCellValue(sheet, cell(col, supervisorInitialsRow), *supervisorInitials)
	f.SetCellValue(sheet, cell(col, dayOfMonthRow), lastDayOfMonth(currentTime(*month)).Format(dayOfMonthFormat))
}

func currentMonthCol(f *excelize.File) (string, error) {
	currentMonthName := currentTime(*month).Format(monthFormat)
	for i := 0; i < months; i++ {
		col := string(monthCol(i))
		monthNameCell := cell(col, monthRow)
		monthName, err := f.GetCellValue(sheet, monthNameCell)
		if err != nil {
			return "", fmt.Errorf("failed to read month name: %v", err)
		}

		clearInitials(f, col)
		if monthName == currentMonthName {
			return col, nil
		}
	}
	return "", errors.New("current month col not found")
}

func cell(col string, row int) string {
	return fmt.Sprintf("%s%d", col, row)
}

func updateLeaveDays(f *excelize.File, col string) {
	leaveCell := cell(col, leaveRow)
	f.SetCellValue(sheet, leaveCell, *leave)
}

func updateDaysEarned(f *excelize.File, col string) error {
	daysEarnedCell := cell(col, daysEarnedRow)
	f.SetCellValue(sheet, daysEarnedCell, daysEarned)

	var res string
	var err error

	startingBalanceCell := cell(col, startingBalanceRow)
	res, err = f.CalcCellValue(sheet, startingBalanceCell)
	if err != nil {
		return fmt.Errorf("failed to starting balance days: %v", err)
	}
	f.SetCellValue(sheet, startingBalanceCell, res)

	if *leave > 0 {
		updateLeaveDays(f, col)
	}

	newBalanceCell := cell(col, newBalanceRow)
	res, err = f.CalcCellValue(sheet, newBalanceCell)
	if err != nil {
		return fmt.Errorf("failed to update days earned: %v", err)
	}
	f.SetCellValue(sheet, newBalanceCell, res)

	return nil
}

func updateCol(f *excelize.File, col string) error {

	updateInitials(f, col)
	return updateDaysEarned(f, col)
}

func updateValues(f *excelize.File, submissionDate string) error {
	f.SetCellValue(sheet, submitterNameCell, fmt.Sprintf("Name: %s", *submitter))
	f.SetCellValue(sheet, dateCell, fmt.Sprintf("Date: %s", submissionDate))
	col, err := currentMonthCol(f)
	if err != nil {
		return fmt.Errorf("failed to update values: %v", err)
	}
	return updateCol(f, col)
}

func main() {
	flag.Parse()

	if err := statFile(outputDir); err != nil {
		fmt.Printf("output dir %s unreadable: %v\n", outputDir, err)
		os.Exit(1)
	}

	prevSubmissionDate := lastDayOfMonth(currentTime(*month - 1))
	submissionDate := lastDayOfMonth(currentTime(*month))
	prevTemplate := genTemplateName(prevSubmissionDate)
	var templateInUse string

	if err := statFile(prevTemplate); err != nil {
		templateInUse = *template
	} else {
		templateInUse = prevTemplate
	}

	f, err := openSheet(templateInUse)
	if err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}

	sheet = f.GetSheetName(*sheetIndex)
	if err := updateValues(f, submissionDate.Format(dateFormat)); err != nil {
		fmt.Printf("failed to update values: %v\n", err)
	}
	if err := saveSheet(f, submissionDate); err != nil {
		fmt.Printf("failed to create updated timesheet: %v\n", err)
	}

}
