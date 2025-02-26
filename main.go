package main

import (
	"fmt"
	"strconv"
	"sort"
	
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
		if (s.quiz+s.midsem+s.labtest+s.weekly_labs)-s.pre_compres > 0.01 {
			is_valid = false
			reasons += "sum of quiz, midsem, labtest, and weekly_labs does not equal pre_compres; "
			fmt.Println(s.quiz + s.midsem + s.labtest + s.weekly_labs)
		}
		// Verify sum of all components
		if (s.quiz+s.midsem+s.labtest+s.weekly_labs+s.compres)-s.total > 0.01 {
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

	fmt.Println("Invalid Student Data Count :", count_invalid)

	// Calculating averages for all components and branch-wise total averages
	validCount := len(rows)-1 - count_invalid
	if validCount <= 0 {
		fmt.Println("No valid student data to compute averages")
	} else {
		var sumQuiz, sumMidsem, sumLabtest, sumWeekly, sumPreCompres, sumCompres, sumTotal float32
		branchSums := make(map[string]float32)
		branchCounts := make(map[string]int)
		for _, s := range students {
			if !s.is_valid {
				continue
			}
			sumQuiz += s.quiz
			sumMidsem += s.midsem
			sumLabtest += s.labtest
			sumWeekly += s.weekly_labs
			sumPreCompres += s.pre_compres
			sumCompres += s.compres
			sumTotal += s.total

			branchSums[s.branch] += s.total
			branchCounts[s.branch]++
		}
		
		// Ranking top 3 students for each component based on their marks
		type RankEntry struct {
			emplid int
			mark   float32
		}
		
		// Helper function to rank components
		rankComponent := func(componentName string, getMark func(s Student) float32) {
			var entries []RankEntry
			for _, s := range students {
				if s.is_valid {
					entries = append(entries, RankEntry{emplid: s.emplid, mark: getMark(s)})
				}
			}
			// sort descending by mark
			sort.Slice(entries, func(i, j int) bool {
				return entries[i].mark > entries[j].mark
			})
			fmt.Printf("Top 3 for %s:\n", componentName)
			ranks := []string{"1st", "2nd", "3rd"}
			for i := 0; i < len(entries) && i < 3; i++ {
				fmt.Printf("Emplid: %d, Marks: %.2f, Rank: %s\n", entries[i].emplid, entries[i].mark, ranks[i])
			}
			fmt.Println()
		}
		fmt.Println("\n========== Component Wise Top 3==========\n")
		
		rankComponent("Quiz", func(s Student) float32 { return s.quiz })
		rankComponent("Midsem", func(s Student) float32 { return s.midsem })
		rankComponent("Labtest", func(s Student) float32 { return s.labtest })
		rankComponent("Weekly Labs", func(s Student) float32 { return s.weekly_labs })
		rankComponent("Pre-Compres", func(s Student) float32 { return s.pre_compres })
		rankComponent("Compres", func(s Student) float32 { return s.compres })
		rankComponent("Total", func(s Student) float32 { return s.total })
		fmt.Println("\n========== General Averages ==========\n")
		fmt.Printf("Average Quiz: %.2f\n", sumQuiz/float32(validCount))
		fmt.Printf("Average Midsem: %.2f\n", sumMidsem/float32(validCount))
		fmt.Printf("Average Labtest: %.2f\n", sumLabtest/float32(validCount))
		fmt.Printf("Average Weekly Labs: %.2f\n", sumWeekly/float32(validCount))
		fmt.Printf("Average Pre-Compres: %.2f\n", sumPreCompres/float32(validCount))
		fmt.Printf("Average Compres: %.2f\n", sumCompres/float32(validCount))
		fmt.Printf("Overall Total Average: %.2f\n", sumTotal/float32(validCount))

		fmt.Println("\n========== Branchwise Total Averages ==========\n")
		for branch, sum := range branchSums {
			avg := sum / float32(branchCounts[branch])
			fmt.Printf("Branch %s Total Average: %.2f\n", branch, avg)
		}
	}
}
