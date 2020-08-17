// 由res2go自动生成。
// 在这里写你的事件。

package main

import (
	"encoding/json"
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"github.com/ying32/govcl/vcl"
	"math"
	"reflect"
	"regexp"
	"strconv"
	"strings"
)

type RowIdxData struct {
	userNameIdx         int // 姓名
	jobTitleIdx         int // 职位
	startIdx            int // 月份开始时间
	endIdx              int // 月份结束
	officialSalaryIdx   int // 正式工资
	internshipSalaryIdx int // 实习工资
	attendanceDays      int // 应出勤天数
}

type TempDataRow struct {
	jobTitle                string
	userName                string
	officialSalary          float64 // 正式工资
	internshipSalary        float64 // 实习工资
	salary                  float64 // 月工资
	baseSalary              float64 // 基本工资 2000
	attendanceDays          float64 // 应出勤天数
	officialDays            float64 // 正式出勤天数
	internshipDays          float64 // 实习天数
	absentDays              float64 // 缺勤天数
	absenceDeduction        float64 // 缺勤扣款
	officePersonalLeave     float64 // 转正事假
	internshipPersonalLeave float64 // 实习事假
	personalLeaveDeduction  float64 // 事假扣款
	officeSickLeave         float64 // 转正病假
	internshipSickLeave     float64 // 实习病假
	sickLeaveDeduction      float64 // 病假扣款
	late                    int     // 迟到
	lateCount               int     // 迟到次数
	lateDeduction           float64 // 迟到扣款
	upUnSignIn              int     // 上班未打卡
	upUnSignInDeduction     float64 // 上班未打卡扣款
	downUnSignIn            int     // 下班未打卡
	downUnSignInDeduction   float64 // 下班未打卡扣款
	unSignInDeduction       float64 // 未打卡扣款
	attendanceAward         float64 // 全勤奖
}

type OutDataRow struct {
	idx                    string // 序号
	userName               string // 姓名
	attendanceDays         string // 应出勤天数
	realAttendanceDays     string // 实际出勤天数
	baseSalary             string // 基本工资
	wageJobs               string // 岗位工资 月工资-基本工资-绩效工资
	performancePay         string // 绩效工资 月工资*0.2
	realSalary             string // 实际工资
	absentDays             string // 缺勤天数
	personalLeave          string // 事假小时
	sickLeave              string // 病假小时
	late                   string // 迟到小时
	foodAllowance          string // 餐补
	attendanceAward        string // 全勤奖
	otherSubsidy           string // 其他补助
	socialSecuritySubsidy  string // 社保补助
	absenceDeduction       string // 缺勤扣款
	sickLeaveDeduction     string // 病假扣款
	personalLeaveDeduction string // 事假扣款
	stayDeduction          string // 住宿扣款
	unSignInDeduction      string // 未打卡扣款
	lateDeduction          string // 迟到扣款
	payable                string // 应发
	socialSecurityCompany  string // 公司社保扣款
	socialSecurityPersona  string // 个人社保扣款
	payable2               string // 应发
	remark                 string // 备注
}

const EIGHT_F = float64(8)
const TWO_F = float64(2)
const HEAD_STYLE_JSON = "{\"Border\":{\"Left\":\"thin\",\"LeftColor\":\"\",\"Right\":\"thin\",\"RightColor\":\"\",\"Top\":\"thin\",\"TopColor\":\"\",\"Bottom\":\"thin\",\"BottomColor\":\"\"},\"Fill\":{\"PatternType\":\"solid\",\"BgColor\":\"\",\"FgColor\":\"FF8EB4E3\"},\"Font\":{\"Size\":11,\"Name\":\"宋体\",\"Family\":0,\"Charset\":134,\"Color\":\"FF000000\",\"Bold\":false,\"Italic\":false,\"Underline\":false,\"Strike\":false},\"ApplyBorder\":true,\"ApplyFill\":true,\"ApplyFont\":true,\"ApplyAlignment\":false,\"Alignment\":{\"Horizontal\":\"\",\"Indent\":0,\"ShrinkToFit\":false,\"TextRotation\":0,\"Vertical\":\"center\",\"WrapText\":false},\"NamedStyleIndex\":0}"
const BODY_STYLE_JSON = "{\"Border\":{\"Left\":\"thin\",\"LeftColor\":\"\",\"Right\":\"thin\",\"RightColor\":\"\",\"Top\":\"thin\",\"TopColor\":\"\",\"Bottom\":\"thin\",\"BottomColor\":\"\"},\"Fill\":{\"PatternType\":\"none\",\"BgColor\":\"\",\"FgColor\":\"\"},\"Font\":{\"Size\":11,\"Name\":\"宋体\",\"Family\":0,\"Charset\":134,\"Color\":\"FF000000\",\"Bold\":false,\"Italic\":false,\"Underline\":false,\"Strike\":false},\"ApplyBorder\":true,\"ApplyFill\":false,\"ApplyFont\":true,\"ApplyAlignment\":false,\"Alignment\":{\"Horizontal\":\"\",\"Indent\":0,\"ShrinkToFit\":false,\"TextRotation\":0,\"Vertical\":\"center\",\"WrapText\":false},\"NamedStyleIndex\":0}"

var personalLeaveExp, _ = regexp.Compile("([A-Z])事(\\d+\\.\\d+)")
var sickLeaveExp, _ = regexp.Compile("([A-Z])病(\\d+\\.\\d+)")
var lateExp, _ = regexp.Compile("([A-Z])迟(\\d+)分")
var upUnSignInExp, _ = regexp.Compile("上班未打")
var downUnSignInExp, _ = regexp.Compile("下班未打")

var bodyStyle = xlsx.NewStyle()
var headStyle = xlsx.NewStyle()

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
	_ = json.Unmarshal([]byte(BODY_STYLE_JSON), bodyStyle)
	_ = json.Unmarshal([]byte(HEAD_STYLE_JSON), headStyle)
	wb, err := xlsx.OpenFile(strFileName)
	if err != nil {
		logToMemoLn("打开excel失败")
		processError(err)
		return
	}
	for i, sheet := range wb.Sheets {
		fmt.Println(i, sheet.Name)
		logToMemoLn("开始处理" + sheet.Name)
		if sheet.MaxRow < 3 {
			logToMemoLn(sheet.Name + "无数据，跳过")
			continue
		}
		rowIdxData := buildRowIdxData(sheet)
		processRowData(rowIdxData, sheet)
	}
	logToMemoLn("处理完成")
	Form1.StartCalc.SetEnabled(true)
}

func buildRowIdxData(sheet *xlsx.Sheet) *RowIdxData {
	// fmt.Println(sheet.MaxCol)
	var rowIdxData = &RowIdxData{}
	row, err := sheet.Row(1)
	if err != nil {
		logToMemoLn("excel数据不正确，中止计算。")
		processError(err)
		return rowIdxData
	}
	for i := 0; i < sheet.MaxCol; i++ {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			logToMemoLn("excel数据为空，中止计算。")
			processError(err)
			break
		}
		if strings.Contains(str, "姓名") {
			rowIdxData.userNameIdx = i
			rowIdxData.startIdx = i + 1
		}
		if strings.Contains(str, "职位") {
			rowIdxData.jobTitleIdx = i
		}
		if strings.Contains(str, "迟到") {
			rowIdxData.endIdx = i - 1
		}
		if strings.Contains(str, "实习基数") {
			rowIdxData.internshipSalaryIdx = i
		}
		if strings.Contains(str, "转正基数") {
			rowIdxData.officialSalaryIdx = i
		}
		if strings.Contains(str, "应出勤") {
			rowIdxData.attendanceDays = i
		}
	}
	return rowIdxData
}

func processRowData(rowIdxData *RowIdxData, sheet *xlsx.Sheet) {
	wb := xlsx.NewFile()
	newSheet, err := wb.AddSheet("工资表")
	if err != nil {
		logToMemoLn("创建输出excel失败")
		processError(err)
		return
	}
	createHeadRow(newSheet)
	for rowIdx := 2; rowIdx < sheet.MaxRow; rowIdx++ {
		tempDataRow := buildTempData(rowIdxData, sheet, rowIdx)
		var outData = OutDataRow{
			idx:                    strconv.Itoa(rowIdx - 1),
			userName:               tempDataRow.userName,
			attendanceDays:         fmt.Sprintf("%.0f", tempDataRow.attendanceDays),
			realAttendanceDays:     fmt.Sprintf("%.0f", tempDataRow.internshipDays+tempDataRow.officialDays),
			baseSalary:             "2000.00",
			wageJobs:               fmt.Sprintf("%.2f", tempDataRow.officialSalary-2000-(tempDataRow.officialSalary*0.2)),
			performancePay:         fmt.Sprintf("%.2f", tempDataRow.officialSalary*0.2),
			realSalary:             fmt.Sprintf("%.2f", tempDataRow.officialSalary),
			absentDays:             fmt.Sprintf("%.0f", tempDataRow.absentDays),
			personalLeave:          fmt.Sprintf("%.0f", tempDataRow.officePersonalLeave+tempDataRow.internshipPersonalLeave),
			sickLeave:              fmt.Sprintf("%.0f", tempDataRow.officeSickLeave+tempDataRow.internshipSickLeave),
			late:                   strconv.Itoa(tempDataRow.late),
			foodAllowance:          "",
			attendanceAward:        fmt.Sprintf("%.2f", tempDataRow.attendanceAward),
			otherSubsidy:           "",
			socialSecuritySubsidy:  "",
			absenceDeduction:       "",
			sickLeaveDeduction:     fmt.Sprintf("%.2f", tempDataRow.sickLeaveDeduction),
			personalLeaveDeduction: fmt.Sprintf("%.2f", tempDataRow.personalLeaveDeduction),
			stayDeduction:          "",
			unSignInDeduction:      fmt.Sprintf("%.2f", tempDataRow.unSignInDeduction),
			lateDeduction:          fmt.Sprintf("%.2f", tempDataRow.lateDeduction),
			payable:                fmt.Sprintf("%.2f", tempDataRow.salary),
			socialSecurityCompany:  "",
			socialSecurityPersona:  "",
			payable2:               fmt.Sprintf("%.2f", tempDataRow.salary),
			remark:                 "",
		}
		createRow(newSheet, outData, bodyStyle)
	}
	outFileName := Form1.SaveExcelFileEdit.Text()
	if "" == outFileName || len(outFileName) == 0 {
		outFileName = Form1.SelectExcelFileEdit.Text() + "_OUT.xlsx"
	}
	err = wb.Save(outFileName)
	if err != nil {
		logToMemoLn("保存输出excel失败")
		processError(err)
		return
	}
	logToMemoLn("生成excel文件成功，文件保存在：" + outFileName)
}

func buildTempData(rowIdxData *RowIdxData, sheet *xlsx.Sheet, rowIdx int) *TempDataRow {
	var tempData = &TempDataRow{baseSalary: 2000.00}
	row, err := sheet.Row(rowIdx)
	if err != nil {
		logToMemoLn("读取excel行失败")
		processError(err)
		return tempData
	}
	jobTitle, _ := row.GetCell(rowIdxData.jobTitleIdx).FormattedValue()
	tempData.jobTitle = jobTitle
	userName, err := row.GetCell(rowIdxData.userNameIdx).FormattedValue()
	if err != nil {
		logToMemoLn("读取excel姓名失败")
		processError(err)
		return tempData
	}
	tempData.userName = userName
	attendanceDays, err := row.GetCell(rowIdxData.attendanceDays).Float()
	if err != nil {
		logToMemoLn("读取excel应出勤失败")
		processError(err)
		return tempData
	}
	tempData.attendanceDays = attendanceDays
	officialSalary, _ := row.GetCell(rowIdxData.officialSalaryIdx).Float()
	if math.IsNaN(officialSalary) {
		officialSalary = float64(0)
	}
	tempData.officialSalary = officialSalary
	internshipSalary, _ := row.GetCell(rowIdxData.internshipSalaryIdx).Float()
	if math.IsNaN(internshipSalary) {
		internshipSalary = float64(0)
	}
	tempData.internshipSalary = internshipSalary
	if officialSalary == 0 && internshipSalary == 0 {
		logToMemoLn("读取excel失败，转正基数或实习基数必填")
		processError(err)
		return tempData
	}

	// 统计出勤和正式、实习天数
	for i := rowIdxData.startIdx; i <= rowIdxData.endIdx; i++ {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			logToMemoLn("读取excel列失败")
			processError(err)
			return tempData
		}
		if strings.Contains(str, "A") {
			tempData.officialDays++
		}
		if strings.Contains(str, "B") {
			tempData.internshipDays++
		}
		personalLeaveMatched := personalLeaveExp.FindStringSubmatch(str)
		if len(personalLeaveMatched) == 3 {
			personalLeave, err := strconv.ParseFloat(personalLeaveMatched[2], 64)
			if err != nil {
				logToMemoLn("事假数据格式不正确" + str)
				processError(err)
				return tempData
			}
			// 请假0.8 就算作是缺勤一天
			if personalLeave == 0.8 {
				if personalLeaveMatched[1] == "A" {
					tempData.officialDays--
				} else {
					tempData.internshipDays--
				}
			} else {
				if personalLeaveMatched[1] == "A" {
					tempData.officePersonalLeave += personalLeave * 10
				} else {
					tempData.internshipPersonalLeave += personalLeave * 10
				}
			}
		}
		sickLeaveMatched := sickLeaveExp.FindStringSubmatch(str)
		if len(sickLeaveMatched) == 3 {
			sickLeave, err := strconv.ParseFloat(sickLeaveMatched[2], 64)
			if err != nil {
				logToMemoLn("病假数据格式不正确" + str)
				processError(err)
				return tempData
			}
			// 请假0.8 就算作是缺勤一天
			if sickLeave == 0.8 {
				if sickLeaveMatched[1] == "A" {
					tempData.officialDays--
				} else {
					tempData.internshipDays--
				}
			} else {
				if sickLeaveMatched[1] == "A" {
					tempData.officeSickLeave += sickLeave * 10
				} else {
					tempData.internshipSickLeave += sickLeave * 10
				}
			}
		}
		lateMatched := lateExp.FindStringSubmatch(str)
		if len(lateMatched) > 0 {
			late, err := strconv.Atoi(lateMatched[2])
			if err != nil {
				logToMemoLn("迟到数据格式不正确" + str)
				processError(err)
				return tempData
			}
			tempData.lateCount++
			if tempData.lateCount <= 3 {
				// 15分钟内（包含15分钟） 扣20元
				if late <= 15 {
					tempData.lateDeduction += 20
				}
				// 16--30分钟 扣50元
				if late > 15 && late <= 30 {
					tempData.lateDeduction += 50
				}
				// 30分钟以上按事假3小时计算
				if late > 30 {
					if lateMatched[1] == "A" {
						tempData.officePersonalLeave += 3
					} else {
						tempData.internshipPersonalLeave += 3
					}
				}
			} else {
				// 迟到超过三次从第四次起每次100元
				tempData.lateDeduction += 100
			}
			tempData.late += late
		}
		if upUnSignInExp.MatchString(str) {
			tempData.upUnSignInDeduction += 50
			tempData.unSignInDeduction += 50
			tempData.upUnSignIn++
		}
		if downUnSignInExp.MatchString(str) {
			tempData.downUnSignInDeduction += 50
			tempData.unSignInDeduction += 50
			tempData.downUnSignIn++
		}
	}
	tempData.personalLeaveDeduction = 0
	tempData.personalLeaveDeduction += tempData.officialSalary / tempData.attendanceDays / EIGHT_F * tempData.officePersonalLeave
	tempData.personalLeaveDeduction += tempData.internshipSalary / tempData.attendanceDays / EIGHT_F * tempData.internshipPersonalLeave

	tempData.sickLeaveDeduction = 0
	tempData.sickLeaveDeduction += tempData.officialSalary / tempData.attendanceDays / EIGHT_F * tempData.officeSickLeave / TWO_F
	tempData.sickLeaveDeduction += tempData.internshipSalary / tempData.attendanceDays / EIGHT_F * tempData.internshipSickLeave / TWO_F
	// 实习出勤
	a := tempData.internshipSalary / tempData.attendanceDays * tempData.internshipDays
	// 转正出勤
	b := tempData.officialSalary / tempData.attendanceDays * tempData.officialDays

	// 当月实际工资=实习期工资 + 转正工资 -（病假+事假+迟到+未打卡扣款）+全勤奖100（未请假迟到等）
	tempData.salary = a + b - (tempData.personalLeaveDeduction + tempData.sickLeaveDeduction + tempData.lateDeduction + tempData.unSignInDeduction)
	// 缺勤天数
	tempData.absentDays = tempData.attendanceDays - (tempData.officialDays + tempData.internshipDays)
	// 全勤
	if tempData.personalLeaveDeduction == 0 && tempData.sickLeaveDeduction == 0 && tempData.lateDeduction == 0 &&
		tempData.unSignInDeduction == 0 && tempData.absentDays == 0 {
		tempData.attendanceAward = 100
		tempData.salary += 100
	}
	logToMemoLn("计算" + tempData.userName + "完成")
	// fmt.Println(tempData.personalLeaveDeduction, tempData.sickLeaveDeduction, tempData.lateDeduction, tempData.unSignInDeduction)
	return tempData
}

func createHeadRow(newSheet *xlsx.Sheet) {
	var outData = OutDataRow{
		idx:                    "序号",
		userName:               "姓名",
		attendanceDays:         "应出勤",
		realAttendanceDays:     "实际出勤",
		baseSalary:             "基本工资",
		wageJobs:               "岗位工资",
		performancePay:         "绩效工资",
		realSalary:             "月发工资",
		absentDays:             "缺勤（天）",
		personalLeave:          "事假(小时）",
		sickLeave:              "病假（小时）",
		late:                   "迟到",
		foodAllowance:          "饭补",
		attendanceAward:        "全勤奖",
		otherSubsidy:           "其他补款",
		socialSecuritySubsidy:  "社保补助",
		absenceDeduction:       "缺勤扣款",
		sickLeaveDeduction:     "病假扣款",
		personalLeaveDeduction: "事假扣款",
		stayDeduction:          "住宿扣款",
		unSignInDeduction:      "未打卡",
		lateDeduction:          "迟到扣款",
		payable:                "应发合计1",
		socialSecurityCompany:  "保险扣款",
		socialSecurityPersona:  "社保扣款",
		payable2:               "应发合计2",
		remark:                 "备注",
	}

	createRow(newSheet, outData, headStyle)
}

func createRow(newSheet *xlsx.Sheet, outData OutDataRow, style *xlsx.Style) {
	newRow := newSheet.AddRow()
	t := reflect.TypeOf(outData)
	v := reflect.ValueOf(outData)
	for k := 0; k < t.NumField(); k++ {
		cell := newRow.AddCell()
		cell.SetStyle(style)
		cell.SetString(v.Field(k).String())
	}
}

func processError(err error) {
	logToMemoLn(err.Error())
}

func logToMemoLn(str string) {
	go func() {
		vcl.ThreadSync(func() {
			Form1.OutputMemo.Lines().Append(str)
		})
	}()

}
