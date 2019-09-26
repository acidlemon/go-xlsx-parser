package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
	"regexp"
	"strconv"
	"strings"
)

type Styles struct {
	cellStyles []cellXf
	numFmts    map[int]format
}

type cellXf struct {
	fontId   int
	fillId   int
	borderId int
	numFmtId int
	xfId     int

	applyFont         bool
	applyAlignment    bool
	applyBorder       bool
	applyFill         bool
	applyNumberFormat bool
}

type format struct {
	formatCode string
	codeType   FormatCodeType
}

var timePattern *regexp.Regexp = regexp.MustCompile(`^datetime\.(date)?(time)?$`)

type FormatCodeType string

func (f FormatCodeType) IsTime() bool {
	return timePattern.MatchString(string(f))
}

var formats map[int]format

func init() {
	formats = make(map[int]format, 40)
	formats[0x00] = format{`@`, `unicode`}
	formats[0x01] = format{`0`, `int`}
	formats[0x02] = format{`0.00`, `float`}
	formats[0x03] = format{`#,##0`, `float`}
	formats[0x04] = format{`#,##0.00`, `float`}
	formats[0x05] = format{`($#,##0_);($#,##0)`, `float`}
	formats[0x06] = format{`($#,##0_);[RED]($#,##0)`, `float`}
	formats[0x07] = format{`($#,##0.00_);($#,##0.00_)`, `float`}
	formats[0x08] = format{`($#,##0.00_);[RED]($#,##0.00_)`, `float`}
	formats[0x09] = format{`0%`, `int`}
	formats[0x0a] = format{`0.00%`, `float`}
	formats[0x0b] = format{`0.00E+00`, `float`}
	formats[0x0c] = format{`# ?/?`, `float`}
	formats[0x0d] = format{`# ??/??`, `float`}
	formats[0x0e] = format{`m-d-yy`, `datetime.date`}
	formats[0x0f] = format{`d-mmm-yy`, `datetime.date`}
	formats[0x10] = format{`d-mmm`, `datetime.date`}
	formats[0x11] = format{`mmm-yy`, `datetime.date`}
	formats[0x12] = format{`h:mm AM/PM`, `datetime.time`}
	formats[0x13] = format{`h:mm:ss AM/PM`, `datetime.time`}
	formats[0x14] = format{`h:mm`, `datetime.time`}
	formats[0x15] = format{`h:mm:ss`, `datetime.time`}
	formats[0x16] = format{`m-d-yy h:mm`, `datetime.datetime`}

	formats[0x25] = format{`(#,##0_);(#,##0)`, `int`}
	formats[0x26] = format{`(#,##0_);[RED](#,##0)`, `int`}
	formats[0x27] = format{`(#,##0.00);(#,##0.00)`, `float`}
	formats[0x28] = format{`(#,##0.00);[RED](#,##0.00)`, `float`}
	formats[0x29] = format{`_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)`, `float`}
	formats[0x2a] = format{`_($*#,##0_);_($*(#,##0);_(*"-"_);_(@_)`, `float`}
	formats[0x2b] = format{`_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)`, `float`}
	formats[0x2c] = format{`_($*#,##0.00_);_($*(#,##0.00);_(*"-"??_);_(@_)`, `float`}
	formats[0x2d] = format{`mm:ss`, `datetime.timedelta`}
	formats[0x2e] = format{`[h]:mm:ss`, `datetime.timedelta`}
	formats[0x2f] = format{`mm:ss.0`, `datetime.timedelta`}
	formats[0x30] = format{`##0.0E+0`, `float`}
	formats[0x31] = format{`@`, `unicode`}

}

func NewStyles(f *zip.File) *Styles {
	reader, err := f.Open()
	if err != nil {
		log.Fatal(err)
	}
	defer reader.Close()

	// parsing states
	var inCellXfs, inNumFmts bool
	c := cellXf{}
	numfmt := format{}
	numFmtId := 0

	cellStyles := []cellXf{}
	numFmts := map[int]format{}

	d := xml.NewDecoder(reader)
	for {
		token, err := d.Token()
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
			case "cellXfs":
				inCellXfs = true
			case "numFmts":
				inNumFmts = true
			case "xf":
				if inCellXfs {
					for _, a := range se.Attr {
						switch a.Name.Local {
						case "fontId":
							val, err := strconv.Atoi(a.Value)
							if err == nil {
								c.fontId = val
							}

						case "fillId":
							val, err := strconv.Atoi(a.Value)
							if err == nil {
								c.fillId = val
							}

						case "borderId":
							val, err := strconv.Atoi(a.Value)
							if err == nil {
								c.borderId = val
							}

						case "xfId":
							val, err := strconv.Atoi(a.Value)
							if err == nil {
								c.xfId = val
							}

						case "numFmtId":
							val, err := strconv.Atoi(a.Value)
							if err == nil {
								c.numFmtId = val
							}

						case "applyFont":
							val, err := strconv.ParseBool(a.Value)
							if err == nil {
								c.applyFont = val
							}

						case "applyNumberFormat":
							val, err := strconv.ParseBool(a.Value)
							if err == nil {
								c.applyNumberFormat = val
							}

						case "applyAlignment":
							val, err := strconv.ParseBool(a.Value)
							if err == nil {
								c.applyAlignment = val
							}

						case "applyBorder":
							val, err := strconv.ParseBool(a.Value)
							if err == nil {
								c.applyBorder = val
							}

						case "applyFill":
							val, err := strconv.ParseBool(a.Value)
							if err == nil {
								c.applyFill = val
							}
						}
					}
				}

			case "numFmt":
				if inNumFmts {
					for _, a := range se.Attr {
						switch a.Name.Local {
						case "formatCode":
							numfmt.formatCode = a.Value
						case "numFmtId":
							val, err := strconv.Atoi(a.Value)
							if err == nil {
								numFmtId = val
							}
						}
					}
					numfmt.codeType = parseFormatCode(numfmt.formatCode)
				}
			}

		case xml.EndElement:
			switch token.(xml.EndElement).Name.Local {
			case "cellXfs":
				inCellXfs = false
			case "numFmts":
				inNumFmts = false
			case "xf":
				if inCellXfs {
					cellStyles = append(cellStyles, c)
					c = cellXf{}
				}
			case "numFmt":
				numFmts[numFmtId] = numfmt
				numfmt = format{}

			}

		case xml.CharData:
		case xml.Comment:
		case xml.ProcInst:
		case xml.Directive:
			// do nothing!!

		default:
			panic("unknown xml token.")
		}
	}

	return &Styles{
		cellStyles: cellStyles,
		numFmts:    numFmts,
	}

}

func parseFormatCode(code string) FormatCodeType {
	datetime := regexp.MustCompile(`(y|m|d|h|s)`)
	datetimeDate := regexp.MustCompile(`(y|d)`)
	datetimeTime := regexp.MustCompile(`(h|s)`)

	var t string
	switch {
	case strings.Contains(code, ";"):
		t = "unicode"
	case datetime.MatchString(code):
		t = "datetime."

		if datetimeDate.MatchString(code) {
			t += "date"
		}
		if datetimeTime.MatchString(code) {
			t += "time"
		}

		if t == "datetime." {
			t += "date"
		}

	default:
		t = "unicode"
	}

	return FormatCodeType(t)
}

func (s *Styles) CellTypeFromStyle(c cellXf) FormatCodeType {
	if c.numFmtId >= len(formats) {
		return s.numFmts[c.numFmtId].codeType
	}

	return formats[c.numFmtId].codeType
}
