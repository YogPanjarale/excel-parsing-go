package main

import (
	"os"
    "database/sql"
    "flag"
    "fmt"
    "log"
    "strconv"
	"postman.bits/task/reporting"
    _ "github.com/mattn/go-sqlite3"
    "github.com/xuri/excelize/v2"
)

func main() {
    filename := flag.String("filename", "./CSF111_202425_01_GradeBook_stripped.xlsx", "Excel file path to read data from")
    sheetname := flag.String("sheet", "CSF111_202425_01_GradeBook", "Sheet name in the Excel file")
    export_type := flag.String("export","","Export type , i.e --export=json for json report , --export=md for markdown report")
	classno := flag.Int("class",0,"class if mentioned , we generate report only for that classno")
    flag.Parse()
	os.Remove("gradebook.db")
    // Open SQLite database
    db, err := sql.Open("sqlite3", "./gradebook.db")
    if err != nil {
        log.Fatal(err)
    }
    defer db.Close()

    // Create table if it does not exist
    sqlStmt := `
        CREATE TABLE IF NOT EXISTS students (
            sino INTEGER,
            classno INTEGER,
            emplid INTEGER,
            campusid TEXT,
            branch TEXT,
            quiz REAL,
            midsem REAL,
            labtest REAL,
            weekly_labs REAL,
            pre_compres REAL,
            compres REAL,
            total REAL,
            is_valid INTEGER
        );
    `
    _, err = db.Exec(sqlStmt)
    if err != nil {
        log.Fatal(err)
    }

    // Load data from Excel into SQLite
    f, err := excelize.OpenFile(*filename)
    if err != nil {
        fmt.Println(err)
        return
    }
    defer func() {
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()

    rows, err := f.GetRows(*sheetname)
    if err != nil {
        fmt.Println(err)
        return
    }

    // Remove rows missing data
    for i, row := range rows {
        if i == 0 {
            continue
        }
        if row[0] == "" {
            fmt.Println("Removing row ", i+1)
            if err := f.RemoveRow(*sheetname, i+1); err != nil {
                fmt.Println(err)
                return
            }
        }
    }

    // Read updated rows
    rows, err = f.GetRows(*sheetname)
    if err != nil {
        fmt.Println(err)
        return
    }

    for i, row := range rows {
        if i == 0 {
            continue
        }

        // Validate data before inserting
        is_valid := true
        reasons := ""
        quiz, _ := strconv.ParseFloat(row[4], 32)
        midsem, _ := strconv.ParseFloat(row[5], 32)
        labtest, _ := strconv.ParseFloat(row[6], 32)
        weekly_labs, _ := strconv.ParseFloat(row[7], 32)
        pre_compres, _ := strconv.ParseFloat(row[8], 32)
        compres, _ := strconv.ParseFloat(row[9], 32)
        total, _ := strconv.ParseFloat(row[10], 32)

        if quiz < -15 || quiz > 30 {
            is_valid = false
            reasons += "quiz out of range (0-30); "
        }
        if midsem < 0 || midsem > 75 {
            is_valid = false
            reasons += "midsem out of range (0-75); "
        }
        if labtest < 0 || labtest > 60 {
            is_valid = false
            reasons += "labtest out of range (0-60); "
        }
        if weekly_labs < 0 || weekly_labs > 30 {
            is_valid = false
            reasons += "weekly_labs out of range (0-30); "
        }
        if pre_compres < -15 || pre_compres > 195 {
            is_valid = false
            reasons += "pre_compres out of range (0-195); "
        }
        if compres < -15 || compres > 105 {
            is_valid = false
            reasons += "compres out of range (0-105); "
        }
        if total < -15 || total > 300 {
            is_valid = false
            reasons += "total out of range (0-300); "
        }

        if (quiz+midsem+labtest+weekly_labs)-pre_compres > 0.01 {
            is_valid = false
            reasons += "sum of quiz, midsem, labtest, and weekly_labs does not equal pre_compres; "
        }
        if (quiz+midsem+labtest+weekly_labs+compres)-total > 0.01 {
            is_valid = false
            reasons += "sum of all components does not equal total; "
        }

        if !is_valid {
            fmt.Println("Invalid student:", row, "Reasons:", reasons)
        }

		if *classno !=0{
			class_no, _ := strconv.Atoi(row[1])
			if class_no != *classno {
				continue // Skip if classno does not match
			}
		}
        // Insert data into SQLite
        stmt, err := db.Prepare("INSERT INTO students(sino, classno, emplid, campusid, branch, quiz, midsem, labtest, weekly_labs, pre_compres, compres, total, is_valid) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)")
        if err != nil {
            log.Fatal(err)
        }
        defer stmt.Close()

        classno, _ := strconv.Atoi(row[1])
        emplid, _ := strconv.Atoi(row[2])
        branch := row[3][4:6]

        _, err = stmt.Exec(i, classno, emplid, row[3], branch, float32(quiz), float32(midsem), float32(labtest), float32(weekly_labs), float32(pre_compres), float32(compres), float32(total), is_valid)
        if err != nil {
            log.Fatal(err)
        }
    }

	

    // Generating and printing report
    printReport(db)
	if *export_type == "json"{
		reporting.JsonReport(db)
	}
	if *export_type == "md" {
		reporting.MarkdownReport(db)
	}

}
func printReport(db *sql.DB) {
    // Calculate averages for all components
    components := []string{"quiz", "midsem", "labtest", "weekly_labs", "pre_compres", "compres", "total"}
    var sumQuiz, sumMidsem, sumLabtest, sumWeekly, sumPreCompres, sumCompres, sumTotal float64
    var validCount int64

    // Get valid count
    row := db.QueryRow("SELECT COUNT(*) FROM students WHERE is_valid = 1")
    err := row.Scan(&validCount)
    if err != nil {
        log.Fatal(err)
    }

    if validCount <= 0 {
        fmt.Println("No valid student data to compute averages")
        return
    }

    // Calculate sums
    rows, err := db.Query("SELECT quiz, midsem, labtest, weekly_labs, pre_compres, compres, total FROM students WHERE is_valid = 1")
    if err != nil {
        log.Fatal(err)
    }
    defer rows.Close()

    for rows.Next() {
        var quiz, midsem, labtest, weekly_labs, pre_compres, compres, total float64
        err := rows.Scan(&quiz, &midsem, &labtest, &weekly_labs, &pre_compres, &compres, &total)
        if err != nil {
            log.Fatal(err)
        }
        sumQuiz += quiz
        sumMidsem += midsem
        sumLabtest += labtest
        sumWeekly += weekly_labs
        sumPreCompres += pre_compres
        sumCompres += compres
        sumTotal += total
    }

    fmt.Print("\n========== Component Wise Top 3 ==========\n\n")
    rankComponent := func(componentName string) {
        rows, err := db.Query(fmt.Sprintf("SELECT emplid, %s FROM students WHERE is_valid = 1 ORDER BY %s DESC LIMIT 3", componentName, componentName))
        if err != nil {
            log.Fatal(err)
        }
        defer rows.Close()

        fmt.Printf("Top 3 for %s:\n", componentName)
        ranks := []string{"1st", "2nd", "3rd"}
        for i := 0; rows.Next(); i++ {
            var emplid int
            var mark float64
            err := rows.Scan(&emplid, &mark)
            if err != nil {
                log.Fatal(err)
            }
            fmt.Printf("Emplid: %d, Marks: %.2f, Rank: %s\n", emplid, mark, ranks[i])
        }
        fmt.Println()
    }

    for _, component := range components {
        rankComponent(component)
    }

    fmt.Print("\n========== General Averages ==========\n\n")
    fmt.Printf("Average Quiz: %.2f\n", sumQuiz/float64(validCount))
    fmt.Printf("Average Midsem: %.2f\n", sumMidsem/float64(validCount))
    fmt.Printf("Average Labtest: %.2f\n", sumLabtest/float64(validCount))
    fmt.Printf("Average Weekly Labs: %.2f\n", sumWeekly/float64(validCount))
    fmt.Printf("Average Pre-Compres: %.2f\n", sumPreCompres/float64(validCount))
    fmt.Printf("Average Compres: %.2f\n", sumCompres/float64(validCount))
    fmt.Printf("Overall Total Average: %.2f\n", sumTotal/float64(validCount))

    // Get branch-wise total averages
    rows, err = db.Query("SELECT branch, AVG(total) FROM students WHERE is_valid = 1 GROUP BY branch")
    if err != nil {
        log.Fatal(err)
    }
    defer rows.Close()

    fmt.Print("\n========== Branchwise Total Averages ==========\n\n")
    for rows.Next() {
        var branch string
        var avg float64
        err := rows.Scan(&branch, &avg)
        if err != nil {
            log.Fatal(err)
        }
        fmt.Printf("Branch %s Total Average: %.2f\n", branch, avg)
    }
}
