package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

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

func get_cities(sheet string) (x, y []string) {
	f, err := excelize.OpenFile("files/City_of_Kelowna.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	var load_loc []string
	var unload_loc []string
	rows, err := f.GetRows("GLELAN-TOLARM")
	fmt.Println(rows[1][3], rows[1][4])
	for _, row := range rows {
		load_loc = append(load_loc, row[3])
		unload_loc = append(unload_loc, row[4])
	}

	return unique(load_loc), unique(unload_loc)

}

func main() {

	var load_loc, unload_loc = get_cities("GLELAN-TOLARM")

	fmt.Println(load_loc)
	fmt.Println(unload_loc)
}
