package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func find_item(slice []string, val string) (int, bool) {
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

func get_cities(file, sheet string) (x, y []string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	var load_loc []string
	var unload_loc []string
	rows, err := f.GetRows(sheet)
	for _, row := range rows {
		load_loc = append(load_loc, row[3])
		unload_loc = append(unload_loc, row[4])
	}

	return unique(load_loc), unique(unload_loc)

}

func get_activites(file, sheet string) (x map[string][][]string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	activites_map := make(map[string][][]string)

	rows, err := f.GetRows(sheet)

	for _, row := range rows[1:] {
		activites_map[row[2]] = append(activites_map[row[2]], row)

	}

	return activites_map

}

func filter_activites(activites map[string][][]string, cities []string) (x map[string][][]string) {
	activites_map_filtered := make(map[string][][]string)

	for key, rows := range activites {
		for _, row := range rows {
			_, found := find_item(cities, row[4])

			if found {
				activites_map_filtered[key] = append(activites_map_filtered[key], row)
			}
		}
	}

	return activites_map_filtered
}

func main() {
	loc_load, loc_unload := get_cities("files/City_of_Kelowna.xlsx", "GLELAN-TOLARM")
	activities := get_activites("files/Shift_Detail_Report_2021-06-021.xlsx", "Data")

	filtered_load := filter_activites(activities, loc_load)
	filtered_unload := filter_activites(activities, loc_unload)

	fmt.Println(filtered_load, filtered_unload)
}
