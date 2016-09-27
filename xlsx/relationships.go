package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
)

type Relationships map[string]relation

type relation struct {
	Target string
	Type   string
}

func NewRelationships(f *zip.File) Relationships {
	reader, err := f.Open()
	if err != nil {
		log.Fatal(err)
	}
	defer reader.Close()

	rel := Relationships{}

	// parsing states
	var inRels bool

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
			case "Relationships":
				inRels = true
			case "Relationship":
				var id string
				r := relation{}
				if inRels {
					for _, a := range se.Attr {
						switch a.Name.Local {
						case "Id":
							id = a.Value
						case "Target":
							r.Target = a.Value
						case "Type":
							r.Type = a.Value
						}
					}

					if id != "" {
						rel[id] = r
					}
				}
			}

		case xml.EndElement:
			switch token.(xml.EndElement).Name.Local {
			case "Relationships":
				inRels = false
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

	return rel
}
