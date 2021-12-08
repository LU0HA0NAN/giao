```go
package main

import "github.com/lhnonline/giao/xlsx"

// o -> order
// w -> width
// t -> title
type SR struct {
	EventTime        string `giao:"o:4;w:10;t:时间"`
	SearchContent    string `giao:"o:1;w:20;t:内容"`
	SearchResultsYes string `giao:"o:2;w:20;t:是否有搜索结果"`
	AreaCode         string `giao:"o:3;w:20;t:地区码"`
}

func main() {
	var s = make([]SR, 0)

	sr01 := SR{"2021-12-08", "giao", "no", "110110"}
	sr02 := SR{"2021-12-08", "饿了吗饿了吗饿了吗饿了吗", "no", "119119"}
	sr03 := SR{"2021-12-08", "郭老师", "yes", "112114"}
	sr04 := SR{"2021-12-08", "翠翠翠", "no", "911911"}
	s = append(s, sr01, sr02, sr03, sr04)

	where := xlsx.CreateXlsx("/Users/giaogiaogiao/go/src/test/another/sr", "search content", SR{}, s)
	println(where)
}
```