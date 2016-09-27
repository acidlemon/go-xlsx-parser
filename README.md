# go-xlsx-parser
SAX like streaming xlsx parser

## SYNOPSIS

```golang
doc := xlsx.Open("your.xlsx")
parser := &MyParser{}
doc.ParseSheet("Sheet 1", parser)
```

```golang
type MyParser struct{}
func (p *MyParser) ReadLine(rowCount int, columns []interface{}) {
     // write your parser here!
}
```

### Example

```golang
package main

import (
	"fmt"
	"strings"
	"time"

	"github.com/acidlemon/go-xlsx-parser/xlsx"
)

func main() {
	doc := xlsx.Open("xlsx/test.xlsx")
	parser := &MyParser{}

	doc.ParseSheet("シート1", parser)
}

type MyParser struct{}

// see xlsx/parser.go
func (p *MyParser) ReadLine(rowCount int, columns []interface{}) {
	fmt.Printf("Line %d:\n", rowCount)
	data := []string{}
	for _, value := range columns {
		switch v := value.(type) {
		case string:
			data = append(data, `"`+v+`"`)

		case int64:
			data = append(data, fmt.Sprintf("%d", v))

		case float64:
			data = append(data, fmt.Sprintf("%f", v))

		case time.Time:
			data = append(data, v.Format(`"2006-01-02 15:04:05"`))
		}
	}

	fmt.Println(" ", strings.Join(data, `,`))
}

```

Enjoy!

