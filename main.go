package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	sino        int
	classno     int
	emplid      int
	campusid    string
	branch      string
	quiz        float32
	midsem      float32
	labtest     float32
	weekly_labs float32
	pre_compres float32
	compres     float32
	total       float32
	is_valid    bool
}

func main() {
	// TODO : Take the filename and sheetname as input from the user
	filename := "./CSF111_202425_01_GradeBook_stripped.xlsx"
	sheetname := "CSF111_202425_01_GradeBook"

	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Remove rows missing data
	rows, err := f.GetRows(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if row[0] == "" {
			fmt.Println("Removing row ", i+1)
			if err := f.RemoveRow(sheetname, i+1); err != nil {
				fmt.Println(err)
				return
			}
		}
	}
	println("Rows: ", len(rows))
	println("Students: ", len(rows)-1)

	// Read the updated rows
	rows, err = f.GetRows(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	students := make([]Student, len(rows)-1)
	count_invalid := 0
	for i, row := range rows {
		if i == 0 {
			continue
		}

		students[i-1].sino = i
		students[i-1].classno, _ = strconv.Atoi(row[1])
		students[i-1].emplid, _ = strconv.Atoi(row[2])
		students[i-1].campusid = row[3]
		students[i-1].branch = row[3][4:6]
		quiz, _ := strconv.ParseFloat(row[4], 32)
		students[i-1].quiz = float32(quiz)
		midsem, _ := strconv.ParseFloat(row[5], 32)
		students[i-1].midsem = float32(midsem)
		labtest, _ := strconv.ParseFloat(row[6], 32)
		students[i-1].labtest = float32(labtest)
		weekly_labs, _ := strconv.ParseFloat(row[7], 32)
		students[i-1].weekly_labs = float32(weekly_labs)
		pre_compres, _ := strconv.ParseFloat(row[8], 32)
		students[i-1].pre_compres = float32(pre_compres)
		compres, _ := strconv.ParseFloat(row[9], 32)
		students[i-1].compres = float32(compres)
		total, _ := strconv.ParseFloat(row[10], 32)
		students[i-1].total = float32(total)

		// Use local variables for validation with reason logging
		is_valid := true
		reasons := ""
		s := students[i-1]

		if s.quiz < -15 || s.quiz > 30 {
			is_valid = false
			reasons += "quiz out of range (0-30); "
		}
		if s.midsem < 0 || s.midsem > 75 {
			is_valid = false
			reasons += "midsem out of range (0-75); "
		}
		if s.labtest < 0 || s.labtest > 60 {
			is_valid = false
			reasons += "labtest out of range (0-60); "
		}
		if s.weekly_labs < 0 || s.weekly_labs > 30 {
			is_valid = false
			reasons += "weekly_labs out of range (0-30); "
		}
		if s.pre_compres < -15 || s.pre_compres > 195 {
			is_valid = false
			reasons += "pre_compres out of range (0-195); "
		}
		if s.compres < -15 || s.compres > 105 {
			is_valid = false
			reasons += "compres out of range (0-105); "
		}
		if s.total < -15 || s.total > 300 {
			is_valid = false
			reasons += "total out of range (0-300); "
		}
		// Verify sum of pre_compres
		if (s.quiz+s.midsem+s.labtest+s.weekly_labs ) - s.pre_compres > 0.01 {
			is_valid = false
			reasons += "sum of quiz, midsem, labtest, and weekly_labs does not equal pre_compres; "
			fmt.Println(s.quiz+s.midsem+s.labtest+s.weekly_labs)
		}
		// Verify sum of all components
		if (s.quiz+s.midsem+s.labtest+s.weekly_labs+s.compres) - s.total > 0.01{
			is_valid = false
			reasons += "sum of all components does not equal total; "
		}

		// Assign computed validity back to student
		s.is_valid = is_valid
		students[i-1] = s

		if !s.is_valid {
			fmt.Println("Invalid student:", s, "Reasons:", reasons)
			count_invalid++
		}
	}

	fmt.Println("Count :", count_invalid)
}