package document

import (
	"strings"
	"testing"
)

func TestEscape(t *testing.T) {
	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("WriteHead Succeed")
	}

	if err := doc.WriteText(NewText(`~!@#$%^&*()_+-={}|[]\;':",./<>?`)); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteText succeed")
	}

	if err := doc.WriteEndHeadWithText(true, "pages", "Hello World", ""); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteEndHead Succeed")
	}

	if err := doc.SaveAS("escape.doc"); err != nil {
		t.Errorf(err.Error())
	}
}

func TestEscapeCauseCrash(t *testing.T) {
	newText := func(words string) *Text {
		words = strings.Replace(words, "%", "&#37;", -1)
		text := &Text{}
		text.Words = words
		text.Color = "000000"
		text.Size = "19"
		text.IsBold = false
		text.IsCenter = false
		return text
	}

	doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("WriteHead Succeed")
	}

	if err := doc.WriteText(newText(`~!@#$%^&*()_+-={}|[]\;':",./<>?`)); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteText succeed")
	}

	if err := doc.WriteEndHeadWithText(true, "pages", "Hello World", ""); err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteEndHead Succeed")
	}

	// 保存
	if err := doc.SaveAS("escapeCauseCrash.doc"); err != nil {
		t.Errorf(err.Error())
	}
}
