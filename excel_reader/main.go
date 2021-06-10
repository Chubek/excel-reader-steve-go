package main

import (
	"flag"
	"fmt"
	"math"
	"regexp"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type strTuple struct {
	a, b string
}

func zip(a, b []string) ([]strTuple, error) {

	if len(a) != len(b) {
		return nil, fmt.Errorf("zip: arguments must be of same length")
	}

	r := make([]strTuple, len(a), len(a))

	for i, e := range a {
		r[i] = strTuple{e, b[i]}
	}

	return r, nil
}

func findItem(slice []string, val string) (int, bool) {
	for i, item := range slice {
		if item == val {
			return i, true
		}
	}
	return -1, false
}

func unique(intSlice []string) []string {
	keys := make(map[string]bool)
	list := []string{}
	for _, entry := range intSlice {
		if _, value := keys[entry]; !value {
			keys[entry] = true
			list = append(list, entry)
		}
	}
	return list
}

func getCities(file, sheet string) (x, y []string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	var loadLoc []string
	var unloadLoc []string
	rows, err := f.GetRows(sheet)
	for _, row := range rows {
		loadLoc = append(loadLoc, row[3])
		unloadLoc = append(unloadLoc, row[4])
	}

	return unique(loadLoc), unique(unloadLoc)

}

func getActivites(file, sheet string) (x map[string][][]string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	activitesMap := make(map[string][][]string)

	rows, err := f.GetRows(sheet)

	for _, row := range rows[1:] {
		activitesMap[row[2]] = append(activitesMap[row[2]], row)

	}

	return activitesMap

}

func filterActivites(activites map[string][][]string, cities []string) (x map[string][][]string) {
	activitesMapFiltered := make(map[string][][]string)

	for key, rows := range activites {
		for _, row := range rows {
			_, found := findItem(cities, row[4])

			if found {
				activitesMapFiltered[key] = append(activitesMapFiltered[key], row)
			}
		}
	}

	return activitesMapFiltered
}

func replaceAtIndex(input string, replacement string, index int) string {
	return input[:index] + replacement + input[index+1:]
}

func parseTime(inputTime string) (x time.Time, y error) {
	parsedTime, err := time.Parse(time.RFC3339, replaceAtIndex(inputTime, string(':')+string(inputTime[len(inputTime)-2]), len(inputTime)-2))

	if err != nil {
		fmt.Println("There was a problem parsing timee, it must be in the format of the original Excel file.")
	}

	return parsedTime, err
}

func multiplyDuration(duration, factor float64) (x float64) {
	return duration * factor
}

func addTime(timeParsed time.Time, minutes float64) (x time.Time) {
	return timeParsed.Add(time.Duration(timeParsed.Hour()) + time.Duration(timeParsed.Minute())*time.Duration(int(math.Round((minutes)))))
}

func returnLabels(index int) (x []string) {
	retArr := []string{fmt.Sprintf("A%d", index), fmt.Sprintf("B%d", index), fmt.Sprintf("C%d", index)}

	return retArr
}

func createValsAndSaveExcelFile(activitesMapLoaded map[string][][]string, activitesMapUnLoaded map[string][][]string, fileloc string, factor float64) {
	f := excelize.NewFile()

	for key, value := range activitesMapLoaded {
		sheetName := fmt.Sprintf("%s_LOADED", key)
		index := f.NewSheet(sheetName)
		f.SetActiveSheet(index)
		f.SetCellValue(sheetName, "A1", "Time")
		f.SetCellValue(sheetName, "B1", "Minutes")
		f.SetCellValue(sheetName, "C1", "Time Completed")
		total := 0.0
		totalIndex := 3
		for i, val := range value {
			if timeMatch, _ := regexp.MatchString(`\d+-\d+-\d+T\d+:\d+:\d+-\d+`, val[0]); timeMatch == false {
				fmt.Println("Time didn't match regex, continuing...")
				continue
			}

			if floatMatch, _ := regexp.MatchString(`0.\d+`, val[5]); floatMatch == false {
				fmt.Println("Float didn't match regex, continuing...")
				continue
			}

			timeParsed, errTime := parseTime(val[0])
			if errTime != nil {
				fmt.Println("Error parsing time, continuing...")
				continue
			}
			floatVal, errFloat := strconv.ParseFloat(val[5], 64)
			if errFloat != nil {
				fmt.Println("Error parsing float, continuing...")
				continue
			}
			timeMultiplied := multiplyDuration(floatVal, factor)
			timeAdded := addTime(timeParsed, timeMultiplied)

			labels := returnLabels(i + 2)
			values := []string{timeParsed.String(), fmt.Sprintf("%f", timeMultiplied), timeAdded.String()}

			zipped, _ := zip(labels, values)

			for _, v := range zipped {
				f.SetCellValue(sheetName, v.a, v.b)
			}

			total += timeMultiplied
			totalIndex += i
		}
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", totalIndex), "Total")
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", totalIndex), fmt.Sprintf("%f", total))
	}

	for key, value := range activitesMapUnLoaded {
		sheetName := fmt.Sprintf("%s_UNLOADED", key)
		index := f.NewSheet(sheetName)
		f.SetActiveSheet(index)
		f.SetCellValue(sheetName, "A1", "Time")
		f.SetCellValue(sheetName, "B1", "Minutes")
		f.SetCellValue(sheetName, "C1", "Time Completed")
		total := 0.0
		totalIndex := 3
		for i, val := range value {
			if timeMatch, _ := regexp.MatchString(`\d+-\d+-\d+T\d+:\d+:\d+-\d+`, val[0]); timeMatch == false {
				fmt.Println("Time didn't match regex, continuing...")
				continue
			}

			if floatMatch, _ := regexp.MatchString(`0.\d+`, val[5]); floatMatch == false {
				fmt.Println("Float didn't match regex, continuing...")
				continue
			}

			timeParsed, errTime := parseTime(val[0])
			if errTime != nil {
				fmt.Println("Error parsing time, continuing...")
				continue
			}
			floatVal, errFloat := strconv.ParseFloat(val[5], 64)
			if errFloat != nil {
				fmt.Println("Error parsing float, continuing...")
				continue
			}
			timeMultiplied := multiplyDuration(floatVal, factor)
			timeAdded := addTime(timeParsed, timeMultiplied)

			labels := returnLabels(i + 2)
			values := []string{timeParsed.String(), fmt.Sprintf("%f", timeMultiplied), timeAdded.String()}

			zipped, _ := zip(labels, values)

			for _, v := range zipped {
				f.SetCellValue(sheetName, v.a, v.b)
			}

			total += timeMultiplied
			totalIndex += i
		}
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", totalIndex), "Total")
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", totalIndex), fmt.Sprintf("%f", total))
	}

	if err := f.SaveAs(fileloc); err != nil {
		fmt.Println(err)
	}

}

func operateMain(locCity, citySheet, locShift, shiftSheet, locFinal string, factor float64) {
	loadedLoc, unloadedLoc := getCities(locCity, citySheet)
	activities := getActivites(locShift, shiftSheet)
	filteredLoc := filterActivites(activities, loadedLoc)
	filteredUnLoc := filterActivites(activities, unloadedLoc)

	createValsAndSaveExcelFile(filteredLoc, filteredUnLoc, locFinal, factor)

}

func main() {
	locCityPtr := flag.String("locCity", "files/City_of_Kelowna.xlsx", "Location of city Excel file")
	citySheetPtr := flag.String("citySheet", "GLELAN-COMCOM", "Sheet of city Excel file")

	shiftPtr := flag.String("locShift", "files/Shift_Detail_Report_2021-06-021.xlsx", "Location of shift Excel file")
	shiftSheetPtr := flag.String("shiftSheet", "Data", "Sheet of shift Excel file")

	finalFilePtr := flag.String("locFinal", "files/output/finalFile.xlsx", "Sheet of final file")

	factorPtr := flag.Float64("factor", 600.0, "Factor for time multiplication")

	flag.Parse()

	operateMain(*locCityPtr, *citySheetPtr, *shiftPtr, *shiftSheetPtr, *finalFilePtr, *factorPtr)

}
