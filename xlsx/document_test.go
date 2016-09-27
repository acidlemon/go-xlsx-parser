package xlsx

import (
	"reflect"
	"testing"
)

var testXlsx [][]interface{} = [][]interface{}{
	[]interface{}{"番号", "名前", "年齢", "住所", ""},
	[]interface{}{int64(1), "Sebastian", int64(28), "東京都千代田区", ""},
	[]interface{}{int64(2), "Michael", int64(42), "北海道旭川市", ""},
	[]interface{}{int64(3), "Nico", int64(26), "静岡県沼津市", ""},
	[]interface{}{int64(4), "Lewis", int64(26), "高知県高知市", ""},
	[]interface{}{int64(5), "Max", int64(17), "大分県由布市", ""},
	[]interface{}{int64(6), "Kimi", int64(34), "島根県松江市", ""},
}

func TestOpen(t *testing.T) {
	doc := Open("test.xlsx")
	parser := &TestParser{t, 0}
	doc.ParseSheet("シート1", parser)

	if parser.maxRowCount != len(testXlsx) {
		t.Errorf("expected maxRowCount = %d, but got %d", len(testXlsx), parser.maxRowCount)
	}
}

type TestParser struct {
	t           *testing.T
	maxRowCount int
}

func (test *TestParser) ReadLine(rowCount int, data []interface{}) {
	test.maxRowCount = rowCount
	rowIndex := rowCount - 1
	columnIndex := 0
	for _, value := range data {
		expect := testXlsx[rowIndex][columnIndex]
		if value != expect {
			test.t.Errorf("[%d, %d] expect %v(%s), but got %v(%s)",
				rowIndex, columnIndex,
				expect, reflect.TypeOf(expect).Name(),
				value, reflect.TypeOf(value).Name(),
			)
		}
		columnIndex++
	}
}
