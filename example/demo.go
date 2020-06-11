package main

import (
	"fmt"

	"github.com/document"
)

func main() {
	doc := document.NewDoc()

	if err := doc.WriteHead(); err != nil {
		fmt.Println(err)
	}

	if err := doc.WriteTitle(document.NewText("测试文档")); err != nil {
		fmt.Println(err)
	}

	if err := doc.WriteTitle3(document.NewText("                                  ———Web应用扫描")); err != nil {
		fmt.Println(err)
	}
	tableHead := [][]interface{}{
		{document.NewText("部门或型号")},
		{document.NewText("部门:研发;型号:martin;")},
		{document.NewText("监督时间")},
		{document.NewText("2020-06-04")},
	}
	table := [][]*document.TableTD{
		{
			document.NewTableTD([]interface{}{document.NewText("监督内容")}),
			document.NewTableTD([]interface{}{document.NewText("你好吗")}),
		},
		{
			document.NewTableTD([]interface{}{document.NewText("主要问题描述")}),
			document.NewTableTD([]interface{}{document.NewText("哈哈哈，我不好")}),
		},
		{
			document.NewTableTD([]interface{}{document.NewText("监督意见或建议")}),
			document.NewTableTD([]interface{}{document.NewText("今天天气好吗")}),
		},
		{
			document.NewTableTD([]interface{}{document.NewText("监督人员：yuelei")}),
		},
		{
			document.NewTableTD([]interface{}{document.NewText("部门领导：")}),
		},
	}
	// 合并单元格操作
	trSpan := [][]int{
		{0, 3},
		{0, 3},
		{0, 3},
		{4},
		{4},
	}
	// 头表格宽度
	tdw := []int{1687, 2993, 5687, 1693}
	// 单元格宽度
	thw := []int{1687, 2993, 5687, 1693}
	// 单元格高度
	tdh := []int{3, 5, 5, 2, 2}

	tableObj := document.NewTable("", true, table, tableHead, thw, trSpan, tdw, tdh)
	if err := doc.WriteTable(tableObj); err != nil {
		fmt.Println(err)
	}
	// 这一行要加上，结束word
	if err := doc.WriteEndHead(); err != nil {
		fmt.Println(err)
	}

	doc.SaveAS("1111.doc")
}
