package xlsx

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

func CreateXlsx(filename, sheetName string, look, inputDataSet interface{}) string {
	filename = getSuffixedFileName(filename)
	tmpFile := excelize.NewFile()

	//
	sheet := tmpFile.NewSheet(sheetName)
	tmpFile.DeleteSheet("Sheet1")

	//
	descList := getGiaoDescList(look)
	titleSet := getTitleSet(descList)
	widthSet := getWidthSet(descList, nil)

	switch reflect.TypeOf(inputDataSet).Kind() {
	case reflect.Slice:
		dataSet := reflect.ValueOf(inputDataSet)
		tmpFile.SetSheetRow(sheetName, "A1", &titleSet)
		fileNameSet := getFileNameSet(descList)
		tmpFile.InsertRow(sheetName, dataSet.Len()+1)
		for i := 0; i < dataSet.Len(); i++ {
			var rowStuffing = make([]interface{}, 0)
			for _, curFiledName := range fileNameSet {
				// fmt.Print(curFiledName, "==>", dataSet.Index(i).FieldByName(curFiledName), "\t")
				rowStuffing = append(rowStuffing, dataSet.Index(i).FieldByName(curFiledName))
			}
			tmpFile.SetSheetRow(sheetName, "A"+strconv.Itoa(i+2), &rowStuffing)
			println()
		}
	}

	for i := 0; i < len(widthSet); i++ {
		tmpFile.SetColWidth(sheetName, string(rune('A'+i)), string(rune('A'+i+1)), float64(widthSet[i]))
	}

	tmpFile.SetActiveSheet(sheet)
	if err := tmpFile.SaveAs(filename); err != nil {
		fmt.Println(err)
	}

	return filename
}

type GiaoDesc struct {
	O         int    `json:"o"`
	W         int    `json:"w"`
	T         string `json:"t"`
	FiledName string
}

func getGiaoDescList(in interface{}) []GiaoDesc {
	t := reflect.TypeOf(in)
	var ret = make([]GiaoDesc, 0)
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		giao := f.Tag.Get("giao")
		if giao == "" {
			continue
		}
		// todo validate the pattern, and fix it if necessary
		desc := fromString(giao)
		desc.FiledName = f.Name
		ret = append(ret, desc)
	}
	sort.SliceStable(ret, func(i, j int) bool {
		return ret[i].O < ret[j].O
	})
	return ret
}

func getTitleSet(descList []GiaoDesc) []interface{} {
	var ret = make([]interface{}, 0)
	for _, v := range descList {
		if v.T == "" {
			// if there is no title defined use fileName
			ret = append(ret, v.FiledName)
		} else {
			ret = append(ret, v.T)
		}
	}
	// fmt.Println("title:", ret)
	return ret
}

func getFileNameSet(descList []GiaoDesc) []string {
	var ret = make([]string, 0)
	for _, v := range descList {
		ret = append(ret, v.FiledName)
	}
	return ret
}

func getWidthSet(descList []GiaoDesc, inputDataSet interface{}) []int {

	var ret = make([]int, 0)

	// no input
	if inputDataSet == nil {
		// if w != 0, use it
		// if w == 0, use fileName width
		for _, v := range descList {
			if v.W != 0 {
				ret = append(ret, v.W)
			} else {
				ret = append(ret, len(v.FiledName))
			}
		}
	} else {
		// with input
		// if w != 0, use it
		// if w == 0, use max length of input
		for _, v := range descList {
			if v.W != 0 {
				ret = append(ret, v.W)
			} else {
				ret = append(ret, getMaxWidthByFiledName(v.FiledName, inputDataSet, len(v.FiledName)))
			}
		}
	}
	// fmt.Println("width:", ret)
	return ret
}

// getMaxWidthByFiledName
// 如果fileName对应的数据皆为空，那么这里使用defaultWidth, defaultWidth应为当前字段的名字的长度
// 不好用
func getMaxWidthByFiledName(fileName string, inputDataSet interface{}, defaultWidth int) int {
	var maxLen = 0
	switch reflect.TypeOf(inputDataSet).Kind() {
	case reflect.Slice:
		dataSet := reflect.ValueOf(inputDataSet)
		for i := 0; i < dataSet.Len(); i++ {
			tmp := dataSet.Index(i).FieldByName(fileName).String()
			tmpLen := utf8.RuneCountInString(tmp)
			if tmpLen > maxLen {
				maxLen = tmpLen
			}
		}
	}
	if maxLen == 0 {
		maxLen = defaultWidth
	}
	if maxLen < defaultWidth {
		maxLen = defaultWidth
	}
	return maxLen
}

func fromString(desc string) GiaoDesc {
	var t GiaoDesc
	split := strings.Split(desc, ";")
	for _, v := range split {
		i := strings.Split(v, ":")
		switch i[0] {
		case "o":
			{
				o, _ := strconv.Atoi(i[1])
				t.O = o
			}
		case "w":
			{
				w, _ := strconv.Atoi(i[1])
				t.W = w
			}
		case "t":
			t.T = i[1]
		}
	}
	return t
}

func getSuffixedFileName(fileName string) string {
	if strings.HasSuffix(fileName, "xlsx") || strings.HasSuffix(fileName, "xls") {
		return fileName
	}
	return fileName + ".xlsx"
}
