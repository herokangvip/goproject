package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"
)

import (
	"github.com/xuri/excelize/v2"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

type MyWindow struct {
	*walk.MainWindow
	ip      *walk.LineEdit
	country *walk.LineEdit
	area    *walk.LineEdit
	region  *walk.LineEdit
	city    *walk.LineEdit
	isp     *walk.LineEdit

	query *walk.PushButton
}

type IPInfo struct {
	Code int `json:"code"`
	Data IP  `json:"data`
}

type IP struct {
	Country string `json:"country"`
	Area    string `json:"area"`
	Region  string `json:"region"`
	City    string `json:"city"`
	Isp     string `json:"isp"`
}

func main() {
	mw := new(MyWindow)
	if err := (MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "excel工具",
		MinSize:  Size{350, 300},
		Layout:   VBox{},
		Children: []Widget{
			Composite{
				MaxSize: Size{0, 50},
				Layout:  HBox{},
				Children: []Widget{
					Label{Text: "文件全路径（包含文件名): "},
					LineEdit{AssignTo: &mw.ip},
					//PushButton{
					//	AssignTo: &mw.query,
					//	Text:     "运行",
					//},
					Label{Text: "名次变动最小阈值: "},
					LineEdit{AssignTo: &mw.country},
					PushButton{
						AssignTo: &mw.query,
						Text:     "运行",
					},
				},
			},
			//Composite{
			//	MinSize: Size{0, 100},
			//	Layout:  HBox{},
			//	Children: []Widget{
			//		GroupBox{
			//			Title:  "查询结果",
			//			Layout: Grid{Columns: 2},
			//			Children: []Widget{
			//				Label{Text: "国家"},
			//				LineEdit{AssignTo: &mw.country, ReadOnly: true},
			//				Label{Text: "地区"},
			//				LineEdit{AssignTo: &mw.area, ReadOnly: true},
			//				Label{Text: "省"},
			//				LineEdit{AssignTo: &mw.region, ReadOnly: true},
			//				Label{Text: "市"},
			//				LineEdit{AssignTo: &mw.city, ReadOnly: true},
			//				Label{Text: "运营商"},
			//				LineEdit{AssignTo: &mw.isp, ReadOnly: true},
			//			},
			//		},
			//	},
			//},
		},
	}).Create(); err != nil {
		log.Fatalln(err)
	}

	mw.query.Clicked().Attach(func() {
		go func() {
			mw.query.SetText("处理中...")
			mw.query.SetEnabled(false)
			mw.GetIpInfo()
			mw.query.SetText("处理完成")
			mw.query.SetEnabled(true)

		}()
	})

	mw.Run()
}

func (mw *MyWindow) GetIpInfo() {
	mw.clearInfo()
	//ip := net.ParseIP(mw.ip.Text())
	ip := mw.ip.Text()
	limit := mw.country.Text()

	if cmd, err := PathExists(ip); err != nil {
		walk.MsgBox(mw, "文件处理", "您输入的不是有效的路径或excel，请重新输入！", walk.MsgBoxIconWarning)
		return
	} else if cmd {
		if strings.LastIndex(ip, ".xl") <= 0 {
			walk.MsgBox(mw, "文件处理", "您输入的不是excel文件，请重新输入！", walk.MsgBoxIconWarning)
			return
		}
		limitParam, err := strconv.Atoi(limit)

		if err != nil {
			walk.MsgBox(mw, "文件处理", "名次变动最小阈值必须为大于0的整数，请重新输入！", walk.MsgBoxIconWarning)
			return
		}
		if limitParam <= 0 {
			walk.MsgBox(mw, "文件处理", "名次变动最小阈值必须为大于0的整数，请重新输入！", walk.MsgBoxIconWarning)
			return
		}

		tabaoAPI(ip, limit)
	} else {
		walk.MsgBox(mw, "文件处理", "您输入的不是有效的路径，请重新输入！", walk.MsgBoxIconWarning)
		return
	}

	//ipResult := tabaoAPI(ip.String())

	//mw.country.SetText(ipResult.Data.Country)
	//mw.area.SetText(ipResult.Data.Area)
	//mw.region.SetText(ipResult.Data.Region)
	//mw.city.SetText(ipResult.Data.City)
	//mw.isp.SetText(ipResult.Data.Isp)

	walk.MsgBox(mw, "文件处理", "处理完毕!", walk.MsgBoxIconInformation)
}

func (mw *MyWindow) clearInfo() {
	//mw.country.SetText("")
	//mw.area.SetText("")
	//mw.region.SetText("")
	//mw.city.SetText("")
	//mw.isp.SetText("")
}

//func tabaoAPI(ip string) *IPInfo {
//	resp, err := http.Get(fmt.Sprintf("http://ip.taobao.com/service/getIpInfo.php?ip=%s", ip))
//	if err != nil {
//		return nil
//	}
//	defer resp.Body.Close()
//
//	out, err := ioutil.ReadAll(resp.Body)
//	if err != nil {
//		return nil
//	}
//	var result IPInfo
//	if err := json.Unmarshal(out, &result); err != nil {
//		return nil
//	}
//
//	return &result
//}
//
func tabaoAPI(fileName string, limit string) bool {
	limitMax, _ := strconv.Atoi(limit)
	//limitMax := 10
	limitMin := limitMax * -1

	fs := excelize.NewFile()

	//t, _ := time.ParseInLocation("20060102150405", "20100424082959", time.Local)

	//字串
	//fs.SetCellValue("Sheet1","B1","LiuBei")
	//时间
	//fs.SetCellValue("Sheet1","B2",t)
	//整形
	//fs.SetCellValue("Sheet1","B3",11)
	//浮点型
	//fs.SetCellValue("Sheet1","B4",3.1415926)

	//open1,_ :=os.Open("D:\\kang\\sanGuo.xlsx")
	//dextFile1, _ := excelize.OpenReader(open1)
	//dextFile1.SetCellStr("Sheet1", "A3", "===========")
	//open1.Close()

	println("start=====")
	srcFile := fileName
	//srcFile := "D:\\kang\\test.xlsx"
	tail := strings.LastIndex(srcFile, ".")
	temp := srcFile[0:tail]
	temp2 := srcFile[tail:]
	println("========================" + temp)
	println("========================" + temp2)
	now := time.Now()
	fmt.Println(now.Unix())
	destFile := temp + strconv.FormatInt(now.Unix(), 10) + temp2
	//destFile := "D:\\kang\\test001.xlsx"
	//total, err := copyFile(srcFile, destFile)
	//fmt.Println(err)
	//fmt.Println(total)

	f, _ := excelize.OpenFile(srcFile)
	lastWeekMap := map[string]int{}
	currentWeekMap := map[string]int{}
	// 获取工作表中指定单元格的值
	sheetName := f.GetSheetName(0)
	rows, _ := f.GetRows(sheetName)
	for rowNum, row := range rows {
		//if rowNum == 0 {
		//	continue
		//}
		fmt.Printf("rowNum,%d\n", rowNum)
		for celNum, celValue := range row {
			if celNum == 0 {
				lastWeekMap[celValue] = rowNum
				fs.SetCellStr(sheetName, "A"+strconv.Itoa(rowNum+1), celValue)
			}
			if celNum == 1 {
				currentWeekMap[celValue] = rowNum
				fs.SetCellStr(sheetName, "C"+strconv.Itoa(rowNum+1), celValue)
			}
			continue
			//fmt.Printf("celNum,%d\n", celNum)
			//fmt.Printf("celValue,%s\n", celValue)
		}
	}

	//destFile中写入新数据
	//open,_ :=os.Open(destFile)
	//dextFile, _ := excelize.OpenReader(open)
	//dextFile.SetCellStr(sheetName, "E3", "===========")
	redStyle, err := fs.NewStyle(`{"fill":{"type":"pattern","color":["#FF0000"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}
	greenStyle, err := fs.NewStyle(`{"fill":{"type":"pattern","color":["#00CD66"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}
	blueGreenStyle, err := fs.NewStyle(`{"fill":{"type":"pattern","color":["#79CDCD"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}
	//粉色
	pinkStyle, err := fs.NewStyle(`{"fill":{"type":"pattern","color":["#CD919E"],"pattern":1}}`)
	if err != nil {
		fmt.Println(err)
	}

	for key, value := range currentWeekMap {
		mingci, ok := lastWeekMap[key]
		if ok {
			//上周排名中存在，看本周前进多少名
			var newMingci = mingci - value
			if newMingci == 0 {
				//没变化
			} else if newMingci > 0 {
				if newMingci < limitMax {
					continue
				}
				msg := "较上周上升" + strconv.Itoa(newMingci) + "名"
				//上升了
				fmt.Println("上升了=Key:", key, "Value:", value, "auxi:", "C"+strconv.Itoa(value))
				fs.SetCellStr(sheetName, "D"+strconv.Itoa(value+1), msg)
				err = fs.SetCellStyle(sheetName, "C"+strconv.Itoa(value+1), "C"+strconv.Itoa(value+1), pinkStyle)

			} else {
				if newMingci > limitMin {
					continue
				}
				//下降了
				msg := "较上周下降" + strconv.Itoa(newMingci*-1) + "名"
				fmt.Println("下降了=Key:", key, "Value:", value, "auxi:", "C"+strconv.Itoa(value))
				fs.SetCellStr(sheetName, "D"+strconv.Itoa(value+1), msg)
				err = fs.SetCellStyle(sheetName, "C"+strconv.Itoa(value+1), "C"+strconv.Itoa(value+1), greenStyle)

			}
		} else {
			//上周不存在
			fs.SetCellStr(sheetName, "D"+strconv.Itoa(value+1), "新进排名")
			err = fs.SetCellStyle(sheetName, "C"+strconv.Itoa(value+1), "C"+strconv.Itoa(value+1), redStyle)

		}

	}

	//上周掉出前300名的

	for key, value := range lastWeekMap {
		_, ok := currentWeekMap[key]
		if ok {
			//未掉队
		} else {
			//掉队
			fs.SetCellStr(sheetName, "B"+strconv.Itoa(value+1), "掉出前300名")
			err = fs.SetCellStyle(sheetName, "A"+strconv.Itoa(value+1), "A"+strconv.Itoa(value+1), blueGreenStyle)

		}

	}

	if err := fs.SaveAs(destFile); err != nil {
		fmt.Println(err)
	}
	println("end=====")

	return true
}

//PathExists 判断一个文件或文件夹是否存在
//输入文件路径，根据返回的bool值来判断文件或文件夹是否存在
func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}
