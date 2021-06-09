package main

import (
	"fmt"
	"math"
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

func parseTime(inputTime string) (x time.Time) {
	parsedTime, err := time.Parse(time.RFC3339, replaceAtIndex(inputTime, string(':')+string(inputTime[len(inputTime)-2]), len(inputTime)-2))

	if err != nil {
		fmt.Println("There was a problem parsing timee, it must be in the format of the original Excel file.")
	}

	return parsedTime
}

func multiplyDuration(duration, factor float64) (x float64) {
	return duration * factor
}

func addTime(timeParsed time.Time, minutes float64) (x time.Time) {
	return timeParsed.Add(time.Duration(timeParsed.Minute()) * time.Duration(int(math.Round((minutes)))))
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
		total := 0.0
		totalIndex := 0
		for i, val := range value {
			timeParsed := parseTime(val[0])
			floatVal, _ := strconv.ParseFloat(val[2], 64)
			timeMultiplied := multiplyDuration(floatVal, factor)
			timeAdded := addTime(timeParsed, timeMultiplied)

			labels := returnLabels(i)
			values := []string{timeParsed.String(), fmt.Sprintf("%f", timeMultiplied), timeAdded.String()}

			zipped, _ := zip(labels, values)

			for _, v := range zipped {
				f.SetCellValue(sheetName, v.a, v.b)
			}

			total += timeMultiplied
			totalIndex += i

		}

		f.SetCellValue(sheetName, fmt.Sprintf("D%d", totalIndex), fmt.Sprintf("%f", total))

	}

}

func main() {
	citiesLoaded, citiesUnloaded := getCities("files/City_of_Kelowna.xlsx", "GLELAN-COMCOM")
	activites := getActivites("files/Shift_Detail_Report_2021-06-021.xlsx", "Data")
	filteredActivitesLoaded := filterActivites(activites, citiesLoaded)
	filteredActivitesUnloaded := filterActivites(activites, citiesUnloaded)

	for key, value := range filteredActivitesLoaded {
		for _, val := range value {
			fmt.Println(key, "\t", val)
		}
	}

	for key, value := range filteredActivitesUnloaded {
		for _, val := range value {
			fmt.Println(key, "\t", val)
		}
	}

}
