package xlsx

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
)

func FromExcel(file, sheet string, look interface{}, skipFirstLine bool) (interface{}, error) {
	f, err := excelize.OpenFile(getSuffixedFileName(file))
	if err != nil {
		fmt.Println(err)
		return nil, errors.New("打开文档失败：" + err.Error())
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, errors.New("读取Sheet" + sheet + "失败：" + err.Error())
	}

	rowOfSheet := len(rows)
	descList := getGiaoDescList(look)
	fileNameSet := getFileNameSet(descList)

	inType := reflect.TypeOf(look)
	fmt.Println(inType)

	inSliceType := reflect.SliceOf(inType)
	inTypeSlice := reflect.MakeSlice(inSliceType, 0, 0)

	var first = 1
	if skipFirstLine {
		first = 2
	}

	for i := first; i <= rowOfSheet; i++ {
		v := reflect.New(inType).Elem()
		for j, fileName := range fileNameSet {
			axis := string(rune('a'+j)) + strconv.Itoa(i)
			value, _ := f.GetCellValue(sheet, axis)
			switch (v.FieldByName(fileName).Interface()).(type) {
			case string:
				v.FieldByName(fileName).SetString(value)
			case int:
				{
					iValue, err := strconv.Atoi(value)
					if err != nil {
						continue
					}
					v.FieldByName(fileName).SetInt(int64(iValue))
				}
			case float32, float64:
				{
					fValue, err := strconv.ParseFloat(value, 64)
					if err != nil {
						continue
					}
					v.FieldByName(fileName).SetFloat(fValue)
				}
			}
		}
		inTypeSlice = reflect.Append(inTypeSlice, v)
	}
	return inTypeSlice.Interface(), nil
}
