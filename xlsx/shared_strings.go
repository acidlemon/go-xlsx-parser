package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
)

type SharedStrings []string

func NewSharedStrings(f *zip.File) SharedStrings {
	reader, err := f.Open()
	if err != nil {
		log.Fatal(err)
	}
	defer reader.Close()

	ss := SharedStrings{}

	// parsing states
	var inSi, inT bool
	buf := make([]byte, 0, 1024)

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
			switch token.(xml.StartElement).Name.Local {
			case "si":
				inSi = true
			case "t":
				inT = true
			}

		case xml.EndElement:
			switch token.(xml.EndElement).Name.Local {
			case "si":
				ss = append(ss, string(buf))
				buf = make([]byte, 0, 1024)
				inSi = false
			case "t":
				inT = false
			}

		case xml.CharData:
			if inSi && inT {
				buf = append(buf, token.(xml.CharData)...)
			}

		case xml.Comment:
		case xml.ProcInst:
		case xml.Directive:
			// do nothing!!

		default:
			panic("unknown xml token.")
		}
	}

	return ss
}
