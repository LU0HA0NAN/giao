```go
package main

import (
	"encoding/json"
	"fmt"
	"github.com/lhnonline/giao/xlsx"
)

// o -> order
// w -> width
// t -> title
type SR struct {
	EventTime        string  `giao:"o:4;w:10;t:时间"`
	SearchContent    string  `giao:"o:1;w:20;t:内容"`
	SearchResultsYes string  `giao:"o:2;w:20;t:是否有搜索结果"`
	AreaCode         int     `giao:"o:3;w:20;t:地区码"`
	ASD              float64 `giao:"o:5;w:20;t:asd"`
	Ignore           string  // 这个字段没有giao标签，不会写入到excel,
}

func main() {
	var s = make([]SR, 0)

	sr01 := SR{"2021-12-08", "giao", "no", 110110, 12.4, ""}
	sr02 := SR{"2021-12-08", "饿了吗饿了吗饿了吗饿了吗", "no", 119119, 3.14, ""}
	sr03 := SR{"2021-12-08", "郭老师", "yes", 112114, 2.717, "🪡不搓"}
	sr04 := SR{"2021-12-08", "啊啊啊", "no", 911911, 69.411, ""}
	s = append(s, sr01, sr02, sr03, sr04)

	// 创建excel
	where := xlsx.CreateXlsx("/Users/haonan/go/src/giao/sr", "search content", SR{}, s)
	println(where)

	// 从excel中读取数
	slice, _ := xlsx.FromExcel("/Users/haonan/go/src/giao/sr.xlsx", "search content", SR{}, true)
	for _, v := range slice.([]SR) {
		marshal, _ := json.Marshal(v)
		fmt.Println(string(marshal))
	}
}

```