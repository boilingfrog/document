# document

### 前言

在`https://github.com/scbizu/report`中实现xml的基础上，实现了`buffer`操作文件， 支持`office2003`。

### 如何使用

基本的文档使用

````
doc := NewDoc()

	if err := doc.WriteHead(); err != nil {
		t.Errorf(err.Error())
	}

	err := doc.WriteTitle(NewText("Hello World"))
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteTitle3 Succeed")
	}

	err = doc.WriteText(NewText("    现代操作系统为解决信息能独立于进程之外被长期存储引入了文件，文件作为进程创建信息的逻辑单元可被多个进程并发使用。在 UNIX 系统中，操作系统为磁盘上的文本与图像、鼠标与键盘等输入设备及网络交互等 I/O 操作设计了一组通用 API，使他们被处理时均可统一使用字节流方式。换言之，UNIX 系统中除进程之外的一切皆是文件，而 Linux 保持了这一特性。为了便于文件的管理，Linux 还引入了目录（有时亦被称为文件夹）这一概念。目录使文件可被分类管理，且目录的引入使 Linux 的文件系统形成一个层级结构的目录树。清单1所示的是普通 Linux 系统的顶层目录结构，其中 /dev 是存放了设备相关文件的目录。"))
	if err != nil {
		t.Errorf(err.Error())
	} else {
		t.Log("TestWriteText succeed")
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
	if err := doc.WriteEndHeadWithText(true, "text", "", "hello world"); err != nil {
		t.Errorf(err.Error())
	}

	if err := doc.SaveAS("demo.doc"); err != nil {
		t.Errorf(err.Error())
	}
````

doc样式  

![](https://img2020.cnblogs.com/blog/1237626/202006/1237626-20200614002043329-529266588.png)





