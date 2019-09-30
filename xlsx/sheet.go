package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"math"
	"regexp"
	"strconv"
	"time"
)

type Sheet struct {
	decoder       *xml.Decoder
	reader        io.ReadCloser
	sharedStrings SharedStrings
	styles        *Styles
}

type cell struct {
	Generated   bool
	StyleIndex  int
	Type        string
	RefOriginal string
	RefCol      int
	RefRow      int
	Column      int
	Value       interface{}
	Format      FormatCodeType
}

func NewSheet(f *zip.File, ss SharedStrings, styles *Styles) *Sheet {
	reader, err := f.Open()
	if err != nil {
		log.Fatal(err)
	}

	d := xml.NewDecoder(reader)
	return &Sheet{
		decoder:       d,
		reader:        reader,
		sharedStrings: ss,
		styles:        styles,
	}
}

func (s *Sheet) Close() {
	if s.reader != nil {
		s.reader.Close()
	}
}

func (s *Sheet) Parse(handler SheetHandler) {

	// parsing state
	inSheetData := false
	inValue := false
	currentRows := []interface{}{}
	var c *cell
	rowCount := 0

	data := ""

	for {
		token, err := s.decoder.Token()
		if err == io.EOF {
			err = nil
			break
		}
		if err != nil {
			panic(err)
		}

		switch token.(type) {
		case xml.StartElement:
			se := token.(xml.StartElement)
			switch se.Name.Local {
			case "sheetData":
				inSheetData = true
			case "row":
				if inSheetData {
					currentRows = []interface{}{}
				}

			case "c":
				var idx int
				var t, ref string

				for _, a := range se.Attr {
					switch a.Name.Local {
					case "s":
						val, err := strconv.Atoi(a.Value)
						if err == nil {
							idx = val
						}
					case "t":
						t = a.Value
					case "r":
						ref = a.Value
					}
				}

				c = &cell{
					StyleIndex:  idx,
					Type:        t,
					RefOriginal: ref,
					Column:      len(currentRows) + 1,
				}

			case "v":
				inValue = true
			}

		case xml.EndElement:
			switch token.(xml.EndElement).Name.Local {
			case "sheetData":
				inSheetData = false
			case "row":
				if inSheetData {
					rowCount++
					handler.ReadRow(rowCount, currentRows)
				}
			case "c":
				cells := parseRel(*c)
				currentRows = append(currentRows, cells...)

				// まずValueの値を埋める
				value := ""
				if c.Type == "s" {
					idx, err := strconv.Atoi(data)
					if err != nil {
						panic(`expect number, actual "` + data + `", err=` + err.Error())
					} else if idx >= len(s.sharedStrings) {
						panic(`SharedString index out of bounds: ` + data)
					}
					value = s.sharedStrings[idx]
				} else {
					value = data
				}

				style := s.styles.cellStyles[c.StyleIndex] // TODO oob check
				c.Format = s.styles.CellTypeFromStyle(style)
				c.Value = value // ひとまず文字列ベースの値をセット

				if c.Type == "" {
					if value != "" && c.Format != "" && c.Format.IsTime() {
						serial, err := strconv.ParseFloat(value, 64)
						if err == nil {
							// time.Timeをセット
							c.Value = convertSerialTime(serial)
						}
					} else {
						// 数値の可能性があるのでよしなに変換
						c.Value = convertIfLikeNumber(value)
					}
				}

				currentRows = append(currentRows, c.Value)

				data = ""
				c = nil
			case "v":
				inValue = false
			}

		case xml.CharData:
			if inValue {
				data += string(token.(xml.CharData))
			}

		case xml.Comment:
		case xml.ProcInst:
		case xml.Directive:
			// do nothing!!

		default:
			panic("unknown xml token.")
		}
	}

}

var relationParser *regexp.Regexp = regexp.MustCompile(`([A-Z]+)(\d+)`)

func parseRel(c cell) []interface{} {
	submatches := relationParser.FindStringSubmatch(c.RefOriginal)
	col, row := submatches[1], submatches[2]

	v := 0
	i := 0
	for _, char := range []byte(col) {
		s := len(col) - i - 1
		i++
		v += int(char-'A'+1) * int(math.Pow(26, float64(s)))
	}

	c.RefCol = v
	rowInt, err := strconv.Atoi(row)
	if err != nil {
		panic(`invalid ref: ` + c.RefOriginal + `, col=` + col + `, row=` + row)
	}
	c.RefRow = rowInt

	if c.Column > v {
		panic(`Detected smaller index than current cell, something wrong. `)
	}

	padding := []interface{}{}
	colIndex := c.Column
	for colIndex < v {
		padding = append(padding, "")
		colIndex++
	}

	return padding
}

func nextColumn(last string) string {
	if last == "" {
		return "A"
	}

	next := []rune(last)

	for i := range last {
		// sort.Reverseを使うのが面倒なのでインデックス操作で逆順にしてる
		ir := len(last) - i - 1
		c := rune(last[ir])

		if c == 'Z' {
			next[ir] = 'A'
			continue
		}

		next[ir] = c + 1
		return string(next)
	}

	return fmt.Sprintf("A%s", string(next))
}

func convertSerialTime(serial float64) time.Time {
	fromUnixEpoch := (serial - 25569) * 24 * 60 * 60
	return time.Unix(int64(math.Round(fromUnixEpoch)), 0).In(time.Local)
}

func convertIfLikeNumber(str string) interface{} {
	// Try Integer -> Try Float -> It's string.
	i, err := strconv.ParseInt(str, 10, 64)
	if err == nil {
		return i
	}

	f, err := strconv.ParseFloat(str, 64)
	if err == nil {
		if math.Trunc(f) == f {
			return int64(f)
		}
		return f
	}

	return str
}
