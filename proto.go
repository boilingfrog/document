package document

import (
	"bufio"
	"bytes"
)

// 通过buffer实现文件的操作
type Document struct {
	Buffer *bytes.Buffer
	Writer *bufio.Writer
}

//Text include text configuration
type Text struct {
	Words    string `json:"word"`
	Color    string `json:"color"`
	Size     string `json:"size"`
	IsBold   bool   `json:"isBold"`
	IsCenter bool   `json:"isCenter"`
}

//Image include image configuration.
type Image struct {
	//This image will link to ?
	Hyperlink string `json:"hyperlink"`
	//destination of the URI in WORD (where it will go to?)
	URIDist string `json:"uridist"`
	//source of the image
	ImageSrc string `json:"imageSrc"`
	//image height  (pixel)
	Height float64 `json:"height"`
	//image width  (pixel)
	Width float64 `json:"width"`
	//Zoom image     (pixel)  You'd bette not to change this default value
	CoordSizeX int `json:"coordSizeX"`
	//Zoom
	CoordSizeY int `json:"coordSizeY"`
}

//TableTD descripes every block of the table
type TableTD struct {
	//TData refers block's element
	TData []interface{} `json:"tdata"`
	//TDBG refers block's background
	TDBG bool `json:"tdbg"`
}

//Table include table configuration.
type Table struct {
	//Tbname  is the name of the table
	Tbname string `json:"tbname"`
	//Text OR Image in the sanme line
	Inline bool `json:"inline"`
	//Table data except table head
	TableBody [][]*TableTD `json:"tablebody"`
	//Table head data
	TableHead [][]interface{} `json:"tableHead"`
	// NOTE: Because of  the title line ,the Total width is 8380.
	//Table head width,you should  list all width inside the table head          (pixel)
	Thw []int `json:"thw"`
	//Table body width ,you should list all width inside the table body     (pixel)
	Tdw []int `json:"tdw"`
	// table height
	Tdh []int `json:"tdh"`
	///////////////////////////////////////////////////////////
	//you can merge cells use GridSpan ,if you need not ,just set 0.
	GridSpan [][]int `json:"gridSpan"`
	//Thcenter set table head center word
	ThCenter bool `json:"thCenter"`
}
