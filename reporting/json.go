package reporting
import (
    "encoding/json"
    "os"
	"fmt"
	"database/sql"
	"log"
)

type Report struct {
    ComponentWiseTop3 map[string][]RankEntry `json:"componentWiseTop3"`
    GeneralAverages   map[string]float64     `json:"generalAverages"`
    BranchwiseAverages map[string]float64    `json:"branchwiseAverages"`
}

type RankEntry struct {
    Emplid int     `json:"emplid"`
    Mark   float64 `json:"mark"`
    Rank   string  `json:"rank"`
}

func JsonReport(db *sql.DB) {
    report := Report{
        ComponentWiseTop3: make(map[string][]RankEntry),
        GeneralAverages:   make(map[string]float64),
        BranchwiseAverages: make(map[string]float64),
    }

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

    // Populate general averages
    report.GeneralAverages["quiz"] = sumQuiz / float64(validCount)
    report.GeneralAverages["midsem"] = sumMidsem / float64(validCount)
    report.GeneralAverages["labtest"] = sumLabtest / float64(validCount)
    report.GeneralAverages["weekly_labs"] = sumWeekly / float64(validCount)
    report.GeneralAverages["pre_compres"] = sumPreCompres / float64(validCount)
    report.GeneralAverages["compres"] = sumCompres / float64(validCount)
    report.GeneralAverages["total"] = sumTotal / float64(validCount)

    // Rank top 3 students for each component
    ranks := []string{"1st", "2nd", "3rd"}
    for _, component := range components {
        rows, err = db.Query(fmt.Sprintf("SELECT emplid, %s FROM students WHERE is_valid = 1 ORDER BY %s DESC LIMIT 3", component, component))
        if err != nil {
            log.Fatal(err)
        }
        defer rows.Close()

        var entries []RankEntry
        for i := 0; rows.Next(); i++ {
            var emplid int
            var mark float64
            err := rows.Scan(&emplid, &mark)
            if err != nil {
                log.Fatal(err)
            }
            entries = append(entries, RankEntry{Emplid: emplid, Mark: mark, Rank: ranks[i]})
        }
        report.ComponentWiseTop3[component] = entries
    }

    // Get branch-wise total averages
    rows, err = db.Query("SELECT branch, AVG(total) FROM students WHERE is_valid = 1 GROUP BY branch")
    if err != nil {
        log.Fatal(err)
    }
    defer rows.Close()

    for rows.Next() {
        var branch string
        var avg float64
        err := rows.Scan(&branch, &avg)
        if err != nil {
            log.Fatal(err)
        }
        report.BranchwiseAverages[branch] = avg
    }

    // Marshal report to JSON
    jsonReportData, err := json.MarshalIndent(report, "", "  ")
    if err != nil {
        log.Fatal(err)
    }

    // Write JSON to file
    err = os.WriteFile("report.json", jsonReportData, 0644)
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("JSON report generated successfully.")
}
