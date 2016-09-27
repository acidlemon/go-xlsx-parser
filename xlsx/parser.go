package xlsx

func Open(path string) *Document {
	doc := NewDocument(path)

	return doc
}

// SheetHandler will receive string / int64 / float64 / time.Time
type SheetHandler interface {
	ReadLine(rowCount int, columns []interface{})
}

//type CSVWriter struct{}
//func (w *CSVWriter) ReadLine(int, columns []interface{}) {
//
//}
