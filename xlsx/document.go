package xlsx

import (
	"archive/zip"
	"log"
	"regexp"
)

type Document struct {
	zipFile       *zip.ReadCloser
	path          string
	workbook      *Workbook
	sharedStrings SharedStrings
	styles        *Styles
	relationships Relationships
	sheets        map[string]*Sheet
}

func NewDocument(path string) *Document {
	reader, err := zip.OpenReader(path)
	if err != nil {
		log.Fatalln(err)
	}

	var wb *Workbook
	var ss SharedStrings
	var styles *Styles
	var rels Relationships

	for _, f := range reader.File {
		switch {
		case f.Name == "xl/workbook.xml":
			wb = NewWorkbook(f)

		case f.Name == "xl/sharedStrings.xml":
			ss = NewSharedStrings(f)

		case f.Name == "xl/styles.xml":
			styles = NewStyles(f)

		case f.Name == "xl/_rels/workbook.xml.rels":
			rels = NewRelationships(f)
		}
	}

	doc := &Document{
		zipFile:       reader,
		path:          path,
		workbook:      wb,
		sharedStrings: ss,
		styles:        styles,
		relationships: rels,
	}

	// one more try
	sheets := map[string]*Sheet{}
	re := regexp.MustCompile(`xl/worksheets/sheet[\d]+\.xml`)
	for _, f := range reader.File {
		if re.MatchString(f.Name) {
			sheet := NewSheet(f, ss, styles)
			sheetName := doc.lookupSheetName(f.Name)
			sheets[sheetName] = sheet
		}
	}

	doc.sheets = sheets

	return doc
}

func (d *Document) Close() {
	if d.zipFile != nil {
		d.zipFile.Close()
	}
}

func (d *Document) ParseSheet(name string, parser SheetHandler) {
	sheet, ok := d.sheets[name]
	if !ok {
		panic(`no such sheet: ` + name)
	}

	sheet.Parse(parser)
}

func (d *Document) lookupSheetName(path string) string {
	// reverse find relationship
	rid := ""
	for k, v := range d.relationships {
		if "xl/"+v.Target == path {
			rid = k
			break
		}
	}

	// reverse find workbook
	sheetName := ""
	for k, v := range d.workbook.sheets {
		if v == rid {
			sheetName = k
			break
		}
	}
	return sheetName
}
