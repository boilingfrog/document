package document

import (
	"testing"
)

func TestNewdoc(t *testing.T) {
	doc := NewDoc()
	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}

func TestWriteTitle1(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}
	err := doc.WriteTitle1(NewText("Hello World"))
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteTitle1 Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}

func TestWriteTitle2(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}
	err := doc.WriteTitle2(NewText("Hello World"))
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteTitle2 Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}

func TestWriteTitle3(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}
	err := doc.WriteTitle3(NewText("Hello World"))
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteTitle3 Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}
func TestWriteBr(t *testing.T) {
	doc := NewDoc()
	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}
	err := doc.WriteBR()
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteBr Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}

func TestWriteText(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}

	err := doc.WriteText(NewText("Hello World"))
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteText succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}
func TestWriteTable(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}
	// tabletd := NewTableTD([]interface{}{{, }, {{"a"}, {"b"}}, {{"xxx"}, {"yyyy"}}})
	td0 := NewTableTD([]interface{}{"aaa"})
	td1 := NewTableTD([]interface{}{"bbb"})
	td2 := NewTableTD([]interface{}{"a"})
	td3 := NewTableTD([]interface{}{"b"})
	td4 := NewTableTD([]interface{}{"xxx"})
	td5 := NewTableTD([]interface{}{"yyyyy"})
	table := [][]*TableTD{{td0, td1}, {td2, td3}, {td4, td5}}
	head := [][]interface{}{{"Hello"}, {"World"}}
	tdw := []int{4190, 4190, 4190, 4190, 4190, 4190}
	thw := []int{4190, 4190}
	trSpan := [][]int{
		{0, 0},
		{0, 0},
		{0, 0},
	}
	tdh := []int{2, 2, 2}

	tableObj := NewTable("test", false, table, head, thw, trSpan, tdw, tdh)
	err := doc.WriteTable(tableObj)
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteTable Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}
func TestWriteImage(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}

	image1 := NewImage("1.png", "./images/offlineWS-102-risk.png", 140.00, 160.00, "")
	image2 := NewImage("2.png", "./images/offlineWS-102-url.png", 140.00, 160.00, "")
	images := []*Image{image1, image2}

	if err := doc.WriteImage(false, "", images...); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteImage Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}

func TestWriteEndHead(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}

	err := doc.WriteEndHeadWithText(true, "pages", "Hello World", "")
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteEndHead Succeed")
	}

	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
}
