package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
)

type Workbook struct {
	sheets map[string]string
}

func NewWorkbook(f *zip.File) *Workbook {
	reader, err := f.Open()
	if err != nil {
		log.Fatal(err)
	}
	defer reader.Close()

	sheets := map[string]string{}

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
			if token.(xml.StartElement).Name.Local != "sheet" {
				// skip
				continue
			}

			attr := token.(xml.StartElement).Attr
			var name, rid string
			for _, a := range attr {
				switch {
				case a.Name.Local == "name":
					name = a.Value

				case a.Name.Space == "http://schemas.openxmlformats.org/officeDocument/2006/relationships" && a.Name.Local == "id":
					rid = a.Value
				}
			}
			sheets[name] = rid

		case xml.EndElement:
		case xml.CharData:
		case xml.Comment:
		case xml.ProcInst:
		case xml.Directive:
			// do nothing!!

		default:
			panic("unknown xml token.")
		}
	}

	return &Workbook{
		sheets: sheets,
	}
}
