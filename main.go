package main

import (
	"fmt"
	"log"
	"regexp"
	"strconv"

	"github.com/gocarina/gocsv"
	"github.com/xuri/excelize/v2"
)

type commonCouple struct {
	StartRowIndex int
	EndRowIndex   int
}

var mergeRe = regexp.MustCompile(`C(\d+):G(\d+)`)

func getCommonCouples(doc *excelize.File) []commonCouple {
	commonCouples := []commonCouple{}

	merges, err := doc.GetMergeCells("Schedule")
	if err != nil {
		log.Fatal(err)
	}

	for _, merge := range merges {
		submatch := mergeRe.FindStringSubmatch(merge[0])

		if len(submatch) != 0 {
			startRowNum, err := strconv.Atoi(submatch[1])
			if err != nil {
				log.Fatal(err)
			}

			endRowNum, err := strconv.Atoi(submatch[1])
			if err != nil {
				log.Fatal(err)
			}

			commonCouple := commonCouple{
				StartRowIndex: startRowNum - 1,
				EndRowIndex:   endRowNum - 1,
			}

			commonCouples = append(commonCouples, commonCouple)
		}
	}

	return commonCouples
}

func isCommonCoupleRow(commonCouples []commonCouple, rowIndex int) bool {
	for _, commonCouple := range commonCouples {
		if rowIndex >= commonCouple.StartRowIndex && rowIndex <= commonCouple.EndRowIndex {
			return true
		}
	}

	return false
}

type couple struct {
	Subject   string `csv:"Subject"`
	StartDate string `csv:"Start Date"`
	StartTime string `csv:"Start Time"`
	EndDate   string `csv:"End Date"`
	EndTime   string `csv:"End Time"`
}

func convertHourFromOmskToMoscow(s string) string {
	num, err := strconv.Atoi(s)
	if err != nil {
		log.Fatal(err)
	}

	return fmt.Sprint(num - 3)
}

var subjectRe = regexp.MustCompile(`(.*?), (.*?), (.*)`)

var subjectTypeMapping = map[string]string{
	"лек":  "ЛЕКЦИЯ",
	"прак": "ПРАКТИКА",
}

func convertSubjectName(s string) string {
	submatch := subjectRe.FindStringSubmatch(s)
	if len(submatch) == 0 {
		return s
	}

	typeFullName, ok := subjectTypeMapping[submatch[3]]
	if ok {
		return fmt.Sprintf("%s: %s, %s", typeFullName, submatch[1], submatch[2])
	}

	return fmt.Sprintf("%s: %s, %s", submatch[3], submatch[1], submatch[2])
}

var dateRe = regexp.MustCompile(`(\d{2})\.(\d{2}).(\d{4})`)
var timeRe = regexp.MustCompile(`(\d{1,2})\.(\d{2})-(\d{1,2})\.(\d{2})`)

func getCouples(doc *excelize.File) []couple {
	commonCouples := getCommonCouples(doc)

	couples := []couple{}

	rows, err := doc.GetRows("Schedule")
	if err != nil {
		log.Fatal(err)
	}

	var curDate string

	for i := 2; i < len(rows); i++ {
		row := rows[i]

		if row[0] != "" {
			submatch := dateRe.FindStringSubmatch(row[0])
			if len(submatch) == 0 {
				log.Fatal(fmt.Errorf("can't parse date: %s", row[0]))
			}

			curDate = fmt.Sprintf("%s/%s/%s", submatch[2], submatch[1], submatch[3])
		}

		if len(row) < 3 {
			continue
		}

		var date string
		if row[1] != "" {
			date = row[1]
		} else {
			date = "19.45-21.20"
		}

		submatch := timeRe.FindStringSubmatch(date)
		if len(submatch) == 0 {
			log.Fatal(fmt.Errorf("can't parse time: %s %s", curDate, date))
		}

		common := isCommonCoupleRow(commonCouples, i)
		if common {
			couple := couple{
				Subject:   convertSubjectName(row[2]),
				StartDate: curDate,
				StartTime: fmt.Sprintf("%s:%s", convertHourFromOmskToMoscow(submatch[1]), submatch[2]),
				EndDate:   curDate,
				EndTime:   fmt.Sprintf("%s:%s", convertHourFromOmskToMoscow(submatch[3]), submatch[4]),
			}

			couples = append(couples, couple)
		} else {
			if len(row) < 6 || row[5] == "" {
				continue
			}

			couple := couple{
				Subject:   convertSubjectName(row[5]),
				StartDate: curDate,
				StartTime: fmt.Sprintf("%s:%s", convertHourFromOmskToMoscow(submatch[1]), submatch[2]),
				EndDate:   curDate,
				EndTime:   fmt.Sprintf("%s:%s", convertHourFromOmskToMoscow(submatch[3]), submatch[4]),
			}

			couples = append(couples, couple)
		}
	}

	return couples
}

func main() {
	doc, err := excelize.OpenFile("schedule.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	defer func() {
		if err := doc.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	days := getCouples(doc)

	csvContent, err := gocsv.MarshalString(&days)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println(csvContent)
}
