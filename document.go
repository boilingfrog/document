package document

import (
	"bufio"
	"bytes"
	"encoding/base64"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
)

// 初始化
func NewDoc() *Document {
	b := bytes.NewBuffer(make([]byte, 0))
	return &Document{
		buffer: b,
		writer: bufio.NewWriter(b),
	}
}

// SaveAs ...
func (doc *Document) SaveAS(name string) error {
	file, err := os.OpenFile(name, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		return err
	}
	defer file.Close()

	if err := doc.writer.Flush(); err != nil {
		return nil
	}

	_, err = doc.buffer.WriteTo(file)

	return err
}

// SaveTo ...
func (doc *Document) SaveTo(writer io.Writer) error {
	if err := doc.writer.Flush(); err != nil {
		return err
	}

	_, err := doc.buffer.WriteTo(writer)

	return err
}

//WriteHead init the header
func (doc *Document) WriteHead() error {
	_, err := doc.writer.WriteString(XMLHead)
	// color.Blue("[LOG]:WriteHead wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return err
}

func (doc *Document) WriteEndHead() error {
	_, err := doc.writer.WriteString(XMLSectBegin)
	if err != nil {
		return err
	}

	_, err = doc.writer.WriteString(XMLSectEnd)
	if err != nil {
		return err
	}

	_, err = doc.writer.WriteString(XMLEndHead)

	return err
}

func (doc *Document) WriteEndHeadWithText(sethdr bool, ftrmode string, hdr string, ftr string) error {
	_, err := doc.writer.WriteString(XMLSectBegin)
	if err != nil {
		return err
	}

	//set HDR
	if sethdr {
		if err := doc.writehdr(hdr); err != nil {
			return err
		}
	}

	//set FTR
	if ftrmode != "" {
		if err := doc.writeftr(ftrmode, ftr); err != nil {
			return err
		}
	}

	if _, err := doc.writer.WriteString(XMLSectEnd); err != nil {
		return err
	}

	_, err = doc.writer.WriteString(XMLEndHead)

	return err
}

//WriteTitle == 居中大标题
func (doc *Document) WriteTitle(text *Text) error {
	color := text.Color
	word := text.Words
	Title := fmt.Sprintf(XMLTitle, color, word)
	_, err := doc.writer.WriteString(Title)

	return err
}

//WriteTitle1 == 标题1的格式
func (doc *Document) WriteTitle1(text *Text) error {
	color := text.Color
	word := text.Words
	Title1 := fmt.Sprintf(XMLTitle1, color, word)
	_, err := doc.writer.WriteString(Title1)

	return err
}

//WriteTitle2 == 标题2的格式
func (doc *Document) WriteTitle2(text *Text) error {
	color := text.Color
	word := text.Words
	Title2 := fmt.Sprintf(XMLTitle2, color, word)
	_, err := doc.writer.WriteString(Title2)

	return err
}

//WriteTitle2WithGrayBg == 灰色panel背景的标题2
func (doc *Document) WriteTitle2WithGrayBg(text string) error {
	Title2Gray := fmt.Sprintf(XMLTitle2WithGrayBg, text)
	_, err := doc.writer.WriteString(Title2Gray)
	//color.Blue("[LOG]:WriteTitle2WithGrayBg Wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return err
}

//WriteTitle3 == 标题3的格式
func (doc *Document) WriteTitle3(text *Text) error {
	color := text.Color
	word := text.Words
	Title3 := fmt.Sprintf(XMLTitle3, color, word)
	_, err := doc.writer.WriteString(Title3)
	//color.Blue("[LOG]:WriteTitle3 Wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return err
}

//WriteTitle3WithGrayBg == 灰色panel背景的标题3
func (doc *Document) WriteTitle3WithGrayBg(text string) error {
	Title3Gray := fmt.Sprintf(XMLTitle3WithGrayBg, text)
	_, err := doc.writer.WriteString(Title3Gray)
	//color.Blue("[LOG]:WriteTitle2WithGrayBg Wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return err
}

//WriteTitle4 == 标题4的格式
func (doc *Document) WriteTitle4(text *Text) error {
	word := text.Words
	Title4 := fmt.Sprintf(XMLTitle4, word)
	_, err := doc.writer.WriteString(Title4)

	return err
}

//WriteText == 正文的格式
func (doc *Document) WriteText(text *Text) error {
	color := text.Color
	size := text.Size
	word := text.Words
	var Text string
	if text.IsCenter {
		if text.IsBold {
			Text = fmt.Sprintf(XMLCenterBoldText, color, size, size, word)
		} else {
			Text = fmt.Sprintf(XMLCenterText, color, size, size, word)
		}
	} else {
		if text.IsBold {
			Text = fmt.Sprintf(XMLBoldText, color, size, size, word)
		} else {
			Text = fmt.Sprintf(XMLText, color, size, size, word)
		}
	}
	_, err := doc.writer.WriteString(Text)

	//color.Blue("[LOG]:WriteText Wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return err
}

//WriteBR == 换行
func (doc *Document) WriteBR() error {
	_, err := doc.writer.WriteString(XMLBr)

	return err
}

//WriteTable  ==表格的格式
func (doc *Document) WriteTable(table *Table) error {
	XMLTable := bytes.Buffer{}
	tbname := table.Tbname
	inline := table.Inline
	tableBody := table.TableBody
	tableHead := table.TableHead
	thw := table.Thw
	gridSpan := table.GridSpan
	tdw := table.Tdw
	tdh := table.Tdh
	var used bool
	used = false
	//handle TableHead :Split with TableBody
	if tableHead != nil {
		tableHeadXML := fmt.Sprintf(XMLTableHead, tbname)

		if _, err := XMLTable.WriteString(tableHeadXML); err != nil {
			return err
		}

		if _, err := XMLTable.WriteString(XMLTableHeadTR); err != nil {
			return err
		}

		for thindex, rowdata := range tableHead {
			thw := fmt.Sprintf(XMLHeadTableTDBegin, strconv.FormatInt(int64(thw[thindex]), 10))
			if _, err := XMLTable.WriteString(thw); err != nil {
				return err
			}

			if inline {
				var err error
				if table.ThCenter {
					_, err = XMLTable.WriteString(XMLHeadTableTDBegin2C)
				} else {
					_, err = XMLTable.WriteString(XMLHeadTableTDBegin2)
				}
				if err != nil {
					return err
				}
			}
			for _, rowEle := range rowdata {
				var err error
				if !inline {
					if table.ThCenter {
						_, err = XMLTable.WriteString(XMLHeadTableTDBegin2C)
					} else {
						_, err = XMLTable.WriteString(XMLHeadTableTDBegin2)
					}
					if err != nil {
						return err
					}
				}
				if image, ok := rowEle.(*Image); ok {
					//rowEle is a resource
					str, err := writeImageToBuffer(image)
					if err != nil {
						return err
					}
					if _, err := XMLTable.WriteString(str); err != nil {
						return err
					}
				} else if text, ok := rowEle.(*Text); ok {
					//not
					color := text.Color
					size := text.Size
					word := text.Words
					var data string
					if text.IsCenter {
						if text.IsBold {
							data = fmt.Sprintf(XMLHeadtableTDTextBC, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDTextC, color, size, size, word)
						}
					} else {
						if text.IsBold {
							data = fmt.Sprintf(XMLHeadtableTDTextB, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDText, color, size, size, word)
						}
					}
					if _, err := XMLTable.WriteString(data); err != nil {
						return err
					}
				}
				if !inline {
					if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
						return err
					}
				}
			}
			if inline {
				if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
					return err
				}
			}
			if _, err := XMLTable.WriteString(XMLHeadTableTDEnd); err != nil {
				return err
			}
		}
		if _, err := XMLTable.WriteString(XMLTableEndTR); err != nil {
			return err
		}
	} else {
		nohead := fmt.Sprintf(XMLTableNoHead, tbname)
		if _, err := XMLTable.WriteString(nohead); err != nil {
			return err
		}
	}
	//Generate formation
	for k, v := range tableBody {

		XMLTable.WriteString(XMLTableTR)
		// 处理gridSpan
		for kk, vv := range v {
			var gridSpanNum int
			if gridSpan != nil && len(gridSpan) >= k && len(gridSpan) > 0 && len(gridSpan[k]) >= kk && len(gridSpan[k]) > 0 {
				gridSpanNum = gridSpan[k][kk]
			}

			// td bg
			var td string
			if vv.TDBG {
				// Span formation
				td = fmt.Sprintf(XMLTableTD, tdw[kk], "E7E6E6", gridSpanNum)
			} else {
				// Span formation
				td = fmt.Sprintf(XMLTableTD, tdw[kk], "auto", gridSpanNum)
			}
			if _, err := XMLTable.WriteString(td); err != nil {
				return err
			}
			tds := 0

			// vv.TData = append(vv.TData, "")

			for _, vvv := range vv.TData {

				table, ok := vvv.(*Table)
				if !inline && !ok {
					if _, err := XMLTable.WriteString(XMLTableTD2); err != nil {
						return err
					}
				}
				if inline && !ok && tds == 0 {
					if _, err := XMLTable.WriteString(XMLTableTD2); err != nil {
						return err
					}

				}
				//if td is a table
				if ok {

					//end with table
					used = true
					tablestr, err := writeTableToBuffer(table)
					if err != nil {
						return err
					}
					if _, err := XMLTable.WriteString(tablestr); err != nil {
						return err
					}
					// FIXME: magic operation
					if _, err := XMLTable.WriteString(XMLMagicFooter); err != nil {
						return err
					}
					//image or text
				} else {
					if icon, ko := vvv.(*Image); ko {
						if icon.Hyperlink != "" {
							if _, err := XMLTable.WriteString(XMLImageLinkTitle); err != nil {
								return err
							}
						}
						if _, err := XMLTable.WriteString(XMLIcon); err != nil {
							return err
						}
						if icon.Hyperlink != "" {
							if _, err := XMLTable.WriteString(XMLImageLinkEnd); err != nil {
								return err
							}
						}
					} else if text, ko := vvv.(*Text); ko {
						var err error
						if text.IsCenter {
							if text.IsBold {
								_, err = XMLTable.WriteString(XMLHeadtableTDTextBC)
							} else {
								_, err = XMLTable.WriteString(XMLHeadtableTDTextC)
							}
						} else {
							if text.IsBold {
								_, err = XMLTable.WriteString(XMLHeadtableTDTextB)
							} else {
								_, err = XMLTable.WriteString(XMLHeadtableTDText)
							}
						}
						if err != nil {
							return err
						}
					}
					//not end with table
					used = false
					// var next bool
					// if kk < len(vv.TData)-1 {
					// 	_, next = vv.TData[tds+1].(*Table)
					// }

					if !inline {
						if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
							return err
						}
					}
					// else if inline && next {
					// 	XMLTable.WriteString(XMLIMGtail)
					// }
				}
				tds++
			}
			//not end with table
			if inline && !used {
				if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
					return err
				}
				//reset inline flag
				// inline = false
			}
			// 写入高度
			if kk == len(v)-1 && len(tdh) >= k && len(tdh) > 0 {
				for i := 0; i < tdh[k]; i++ {
					if _, err := XMLTable.WriteString(fmt.Sprintf(XMLTableTDHeight)); err != nil {
						return err
					}
				}
			}
			if _, err := XMLTable.WriteString(XMLHeadTableTDEnd); err != nil {
				return err
			}
		}
		if _, err := XMLTable.WriteString(XMLTableEndTR); err != nil {
			return err
		}
	}
	if _, err := XMLTable.WriteString(XMLTableFooter); err != nil {
		return err
	}
	//serialization
	var rows []interface{}

	for _, row := range tableBody {
		for _, rowdata := range row {
			for _, rowEle := range rowdata.TData {
				if _, ok := rowEle.([][][]interface{}); !ok {
					if icon, ok := rowEle.(*Image); ok {
						//图片
						imageSrc := icon.ImageSrc
						bindata, err := getImagedata(imageSrc)
						URI := "wordml://" + icon.URIDist
						if err != nil {
							return err
						}

						if icon.Hyperlink != "" {
							rows = append(rows, icon.Hyperlink, URI, bindata, filepath.Base(imageSrc), URI, filepath.Base(imageSrc))
						} else {
							rows = append(rows, URI, bindata, filepath.Base(imageSrc), URI, filepath.Base(imageSrc))
						}
					} else if text, ok := rowEle.(*Text); ok {
						tColor := text.Color
						tSize := text.Size
						tWord := text.Words
						rows = append(rows, tColor, tSize, tSize, tWord)
					}
				}
			}
		}
	}

	//data fill in

	tabledata := fmt.Sprintf(XMLTable.String(), rows...)

	_, err := doc.writer.WriteString(tabledata)
	return err
}

// WriteImage == 写入图片
func (doc *Document) WriteImage(withtext bool, text string, imagesData ...*Image) error {
	xmlimage := bytes.Buffer{}
	//write fontStyle

	if _, err := xmlimage.WriteString(XMLIMGTitle); err != nil {
		return err
	}
	if withtext {
		//偷个懒  指定为1
		fontStyle := fmt.Sprintf(XMLFontStyle, "1")
		if _, err := xmlimage.WriteString(fontStyle); err != nil {
			return err
		}
	}
	for _, imagedata := range imagesData {
		imageSrc := imagedata.ImageSrc
		URIDist := imagedata.URIDist
		coordsizeX := imagedata.CoordSizeX
		coordsizeY := imagedata.CoordSizeY
		height := imagedata.Height
		width := imagedata.Width
		hyperlink := imagedata.Hyperlink
		//embedding hyperlink
		if hyperlink != "" {
			imageLink := fmt.Sprintf(XMLImageLinkTitle, hyperlink)
			if _, err := xmlimage.WriteString(imageLink); err != nil {
				return err
			}
		}
		bindata, err := getImagedata(imageSrc)
		if err != nil {
			return err
		}
		URI := "wordml://" + URIDist
		imageSec := fmt.Sprintf(XMLImage, URI, bindata, filepath.Base(imageSrc), strconv.FormatFloat(height, 'f', -1, 64),
			strconv.FormatFloat(width, 'f', -1, 64), strconv.Itoa(coordsizeY), strconv.Itoa(coordsizeX), URI, filepath.Base(imageSrc))
		if _, err := xmlimage.WriteString(imageSec); err != nil {
			return err
		}
		//hyper link
		if hyperlink != "" {
			if _, err := xmlimage.WriteString(XMLImageLinkEnd); err != nil {
				return err
			}
		}
	}
	if withtext {
		inlineText := fmt.Sprintf(XMLInlineText, text)
		if _, err := xmlimage.WriteString(inlineText); err != nil {
			return err
		}
	}
	xmlimage.WriteString(XMLIMGtail)
	_, err := doc.writer.WriteString(xmlimage.String())
	return err
}

// writeImageToBuffer write image xml to buffer and return.
func writeImageToBuffer(image *Image) (string, error) {
	ResImage := bytes.Buffer{}
	// xmlimage := bytes.Buffer{}
	// xmlimage.WriteString(XMLIMGTitle)
	if image.Hyperlink != "" {
		if _, err := ResImage.WriteString(XMLImageLinkTitle); err != nil {
			return "", err
		}
	}
	imageSrc := image.ImageSrc
	URI := "wordml://" + image.URIDist

	bindata, err := getImagedata(imageSrc)
	if err != nil {
		return "", err
	}
	imageSec := fmt.Sprintf(XMLIcon, URI, bindata, filepath.Base(imageSrc), URI, filepath.Base(imageSrc))
	if _, err := ResImage.WriteString(imageSec); err != nil {
		return "", err
	}
	if _, err := ResImage.WriteString(XMLImageLinkEnd); err != nil {
		return "", err
	}
	return ResImage.String(), nil
}

//Generate table xml string formation  ~> 用于 表中再次嵌入表格时的填充
func writeTableToBuffer(table *Table) (string, error) {
	tbname := table.Tbname
	tableHead := table.TableHead
	tableBody := table.TableBody
	inline := table.Inline
	thw := table.Thw
	tdw := table.Tdw
	XMLTable := bytes.Buffer{}
	var Bused bool
	Bused = false
	//handle TableHead :Split with TableBody
	if tableHead != nil {
		//表格中的表格为无边框形式
		tableHeadXML := fmt.Sprintf(XMLTableInTableHead, tbname)
		if _, err := XMLTable.WriteString(tableHeadXML); err != nil {
			return "", err
		}
		if _, err := XMLTable.WriteString(XMLTableHeadTR); err != nil {
			return "", err
		}
		for thindex, rowdata := range tableHead {
			thw := fmt.Sprintf(XMLHeadTableTDBegin, strconv.FormatInt(int64(thw[thindex]), 10))
			if _, err := XMLTable.WriteString(thw); err != nil {
				return "", err
			}
			if inline {
				var err error
				if table.ThCenter {
					_, err = XMLTable.WriteString(XMLHeadTableTDBegin2C)
				} else {
					_, err = XMLTable.WriteString(XMLHeadTableTDBegin2)
				}
				if err != nil {
					return "", err
				}
			}
			for _, rowEle := range rowdata {
				if !inline {
					var err error
					if table.ThCenter {
						_, err = XMLTable.WriteString(XMLHeadTableTDBegin2C)
					} else {
						_, err = XMLTable.WriteString(XMLHeadTableTDBegin2)
					}
					if err != nil {
						return "", err
					}
				}
				if image, ok := rowEle.(*Image); ok {
					//rowEle is a resource
					str, err := writeImageToBuffer(image)
					if err != nil {
						return "", err
					}
					if _, err := XMLTable.WriteString(str); err != nil {
						return "", err
					}
				} else if text, ok := rowEle.(*Text); ok {
					//not
					color := text.Color
					size := text.Size
					word := text.Words
					var data string
					if text.IsCenter {
						// println(text.IsCenter)
						if text.IsBold {
							// println(text.IsBold)
							data = fmt.Sprintf(XMLHeadtableTDTextBC, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDTextC, color, size, size, word)
						}
					} else {
						if text.IsBold {
							data = fmt.Sprintf(XMLHeadtableTDTextB, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDText, color, size, size, word)
						}
					}
					if _, err := XMLTable.WriteString(data); err != nil {
						return "", err
					}
				}
				if !inline {
					if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
						return "", err
					}
				}
			}
			if inline {
				if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
					return "", err
				}
			}
			if _, err := XMLTable.WriteString(XMLHeadTableTDEnd); err != nil {
				return "", err
			}
		}
		if _, err := XMLTable.WriteString(XMLTableEndTR); err != nil {
			return "", err
		}
	} else {
		nohead := fmt.Sprintf(XMLTableInTableNoHead, tbname)

		if _, err := XMLTable.WriteString(nohead); err != nil {
			return "", err
		}
	}

	//Generate formation
	for _, v := range tableBody {
		if _, err := XMLTable.WriteString(XMLTableTR); err != nil {
			return "", err
		}

		for kk, vv := range v {

			var ttd string
			if vv.TDBG {
				//fill with gray
				ttd = fmt.Sprintf(XMLTableInTableTD, strconv.FormatInt(int64(tdw[kk]), 10), "E7E6E6")
			} else {
				ttd = fmt.Sprintf(XMLTableInTableTD, strconv.FormatInt(int64(tdw[kk]), 10), "auto")
			}
			if _, err := XMLTable.WriteString(ttd); err != nil {
				return "", err
			}

			tds := 0
			// vv.TData = append(vv.TData, "")
			if inline {
				if _, err := XMLTable.WriteString(XMLTableTD2); err != nil {
					return "", err
				}

			}
			for _, vvv := range vv.TData {
				table, ok := vvv.(*Table)
				if !inline && !ok {
					if _, err := XMLTable.WriteString(XMLTableTD2); err != nil {
						return "", err
					}
				}

				//if td is a table
				if ok {
					//end with table
					Bused = true
					tablestr, err := writeTableToBuffer(table)
					if err != nil {
						return "", err
					}
					if _, err := XMLTable.WriteString(tablestr); err != nil {
						return "", err
					}
					// FIXME: magic operation
					if _, err := XMLTable.WriteString(XMLMagicFooter); err != nil {
						return "", err
					}
					//image or text
				} else {
					if icon, ko := vvv.(*Image); ko {
						if icon.Hyperlink != "" {
							if _, err := XMLTable.WriteString(XMLImageLinkTitle); err != nil {
								return "", err
							}
						}
						if _, err := XMLTable.WriteString(XMLIcon); err != nil {
							return "", err
						}
						if icon.Hyperlink != "" {
							if _, err := XMLTable.WriteString(XMLImageLinkEnd); err != nil {
								return "", err
							}
						}
					} else if text, ko := vvv.(*Text); ko {
						var err error
						if text.IsCenter {
							if text.IsBold {
								_, err = XMLTable.WriteString(XMLHeadtableTDTextBC)
							} else {
								_, err = XMLTable.WriteString(XMLHeadtableTDTextC)
							}
						} else {
							if text.IsBold {
								_, err = XMLTable.WriteString(XMLHeadtableTDTextB)
							} else {
								_, err = XMLTable.WriteString(XMLHeadtableTDText)
							}
						}
						if err != nil {
							return "", err
						}
					}
					//not end with table
					Bused = false
					var next bool
					if tds < len(vv.TData)-1 {
						_, next = vv.TData[tds+1].(*Table)
					}
					var err error
					if !inline {
						_, err = XMLTable.WriteString(XMLIMGtail)
					} else if inline && next {
						_, err = XMLTable.WriteString(XMLIMGtail)
					}
					if err != nil {
						return "", err
					}
				}
				tds++
			}
			//not end with table
			if inline && !Bused {
				if _, err := XMLTable.WriteString(XMLIMGtail); err != nil {
					return "", err
				}
				//reset inline flag
				// inline = false
			}
			if _, err := XMLTable.WriteString(XMLHeadTableTDEnd); err != nil {
				return "", err
			}

		}
		if _, err := XMLTable.WriteString(XMLTableEndTR); err != nil {
			return "", err
		}
	}
	if _, err := XMLTable.WriteString(XMLTableFooter); err != nil {
		return "", err
	}
	//serialization
	var rows []interface{}

	for _, row := range tableBody {
		for _, rowdata := range row {
			for _, rowEle := range rowdata.TData {
				if _, ok := rowEle.([][][]interface{}); !ok {
					if icon, ok := rowEle.(*Image); ok {
						//图片
						imageSrc := icon.ImageSrc
						bindata, err := getImagedata(imageSrc)
						URI := "wordml://" + icon.URIDist
						if err != nil {
							return "", err
						}

						if icon.Hyperlink != "" {
							rows = append(rows, icon.Hyperlink, URI, bindata, filepath.Base(imageSrc), URI, filepath.Base(imageSrc))
						} else {
							rows = append(rows, URI, bindata, filepath.Base(imageSrc), URI, filepath.Base(imageSrc))
						}
					} else if text, ok := rowEle.(*Text); ok {
						tColor := text.Color
						tSize := text.Size
						tWord := text.Words
						rows = append(rows, tColor, tSize, tSize, tWord)
					}
				}
			}
		}
	}

	//data fill in
	tabledata := fmt.Sprintf(XMLTable.String(), rows...)

	return tabledata, nil
}

//get bindata
func getImagedata(src string) (string, error) {
	file, err := os.Open(src)
	if err != nil {
		return "", err
	}
	defer file.Close()
	//Get bindata , encode via Base64
	finfo, _ := file.Stat()
	size := finfo.Size()
	buf := make([]byte, size)
	encoder := bufio.NewReader(file)
	encoder.Read(buf)
	bindata := base64.StdEncoding.EncodeToString(buf)
	return bindata, nil
}

//writehdr ==页眉格式  wrap fucntion
func (doc *Document) writehdr(text string) error {
	hdr := fmt.Sprintf(XMLhdr, text)
	_, err := doc.writer.WriteString(hdr)
	if err != nil {
		return err
	}
	//color.Blue("[LOG]:WriteTitle1 Wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return nil
}

//writeftr == 页脚  wrap function
/////MODE :
//"pages":For page index
//"text" :For footer  text
//others :none
func (doc *Document) writeftr(mode string, text string) error {
	switch mode {
	case "pages":
		if _, err := doc.writer.WriteString(XMLftrPages); err != nil {
			return err
		}
	case "text":
		ftrtext := fmt.Sprintf(XMLftrText, text)
		if _, err := doc.writer.WriteString(ftrtext); err != nil {
			return err
		}
	default:
		return errors.New("Unknown Footer Mode :(")
	}

	// color.Blue("[LOG]:WriteTitle1 Wrote" + strconv.FormatInt(int64(count), 10) + "bytes")
	return nil
}

// if the str is a resource file
// BUG: other resorce can not be imported
func isResource(str string) bool {
	file, err := os.Open(str)
	if err != nil {
		return false
	}
	defer file.Close()
	return true
}

//NewImage init a image with fixed CoordsizeX & CoordsizeY
func NewImage(URIdist string, imageSrc string, height float64, width float64, hyperlink string) *Image {
	img := &Image{}
	img.URIDist = URIdist
	img.ImageSrc = imageSrc
	img.Height = height
	img.Width = width
	img.CoordSizeX = 21600
	img.CoordSizeY = 21600
	img.Hyperlink = hyperlink
	return img
}

//NewTable create a table
func NewTable(tbname string, inline bool, tableBody [][]*TableTD, tableHead [][]interface{}, thw []int, gridSpan [][]int, tdw []int, tdh []int) *Table {
	table := &Table{}
	table.Tbname = tbname
	table.Inline = inline
	table.TableBody = tableBody
	table.TableHead = tableHead
	table.Tdw = tdw
	table.Thw = thw
	table.Tdh = tdh

	table.GridSpan = gridSpan
	table.ThCenter = false
	return table
}

//SetHeadCenter set table head center word
func (tb *Table) SetHeadCenter(center bool) {
	tb.ThCenter = center
}

//NewText create word with default setting
func NewText(words string) *Text {
	words = wordescape(words)
	text := &Text{}
	text.Words = words
	text.Color = "000000"
	text.Size = "19"
	text.IsBold = false
	text.IsCenter = false
	return text
}

//Setcolor Set Text color
func (tx *Text) Setcolor(color string) {
	tx.Color = color
}

//SetSize set text size
func (tx *Text) SetSize(size string) {
	tx.Size = size
}

//SetBold set bold
func (tx *Text) SetBold(bold bool) {
	tx.IsBold = bold
}

//SetCenter set center  text
func (tx *Text) SetCenter(center bool) {
	tx.IsCenter = center
}

//NewTableTD init table td block
func NewTableTD(tdata []interface{}) *TableTD {
	Tabletd := &TableTD{}
	Tabletd.TData = tdata
	Tabletd.TDBG = false
	return Tabletd
}

//SetTableTDBG set block's color with gray(#E7E6E6)
func (tbtd *TableTD) SetTableTDBG() {
	tbtd.TDBG = true
}

//Solve  the  '%' cause (MISSING) crash problem
func wordescape(str string) string {
	return Escape(str)
}
