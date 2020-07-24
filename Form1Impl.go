// 由res2go自动生成。
// 在这里写你的事件。

package main

import (
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"github.com/ying32/govcl/vcl"
	"strconv"
)

type rowData struct {
	idx int  //序号
	position string  //职位
	name string //姓名
	salaryBase1 int //实习工资基数
	salaryBase2 int //转正工资基数
	cells        []*int
	late int     // 迟到
	L string     // 最小值
	A string     // 平均值
	TIME string  // 时间
}

//::private::
type TForm1Fields struct {
}

func (f *TForm1) OnSelectExcelFileBtnClick(sender vcl.IObject) {

}

func (f *TForm1) OnSaveExcelFileBtnClick(sender vcl.IObject) {

}

func (f *TForm1) OnSelectExcelFileActExecute(sender vcl.IObject) {
	Form1.SelectExcelFileDia.SetFilter("excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|全部文件(*.*)|*.*")
	Form1.SelectExcelFileDia.SetDefaultExt("*.xlsx")
	if Form1.SelectExcelFileDia.Execute() {
		strFileName := Form1.SelectExcelFileDia.FileName()
		if strFileName != "" {
			Form1.SelectExcelFileEdit.SetText(strFileName)
			logToMemoLn("选择的文件：\n" + strFileName + "\n")
		}
	}
}

func (f *TForm1) OnSaveExcelFileActExecute(sender vcl.IObject) {
	Form1.SaveExcelFileDia.SetFilter("excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|全部文件(*.*)|*.*")
	Form1.SaveExcelFileDia.SetDefaultExt("*.xlsx")
	if Form1.SaveExcelFileDia.Execute() {
		strFileName := Form1.SaveExcelFileDia.FileName()
		if strFileName != "" {
			Form1.SaveExcelFileEdit.SetText(strFileName)
			logToMemoLn("保存新文件到：\n" + strFileName + "\n")
		}
	}
}

func (f *TForm1) OnStartCalcClick(sender vcl.IObject) {
	Form1.StartCalc.SetEnabled(false)
	strFileName := Form1.SelectExcelFileEdit.Text()
	fmt.Println(strFileName)
	wb, err := xlsx.OpenFile(strFileName)
	if err != nil {
		panic(err)
	}
	for i, sheet := range wb.Sheets {
		fmt.Println(i, sheet.Name)
		logToMemoLn("开始处理" + sheet.Name)
		if sheet.MaxRow < 3 {
			logToMemoLn(sheet.Name + "无数据，跳过")
			continue
		}
		err := sheet.ForEachRow(rowVisitor)
		if err != nil {
			panic("处理sheet出错")
		}
	}
	logToMemoLn("处理完成")
	Form1.StartCalc.SetEnabled(true)
}

func rowVisitor(row *xlsx.Row) error {
	// fmt.Println(row)
	/*err := row.ForEachCell(cellVisitor)
	if err != nil {
		panic("处理row出错")
	}*/
	return nil
}

func cellVisitor(c *xlsx.Cell) error {
	value, err := c.FormattedValue()
	if err != nil {
		fmt.Println(err.Error())
	} else {
		fmt.Println("Cell value:", value)
	}
	return err
}

func logIntToMemoLn(num int) {
	Form1.OutputMemo.Lines().Append(strconv.Itoa(num))
}

func logToMemoLn(str string) {
	Form1.OutputMemo.Lines().Append(str)
}
