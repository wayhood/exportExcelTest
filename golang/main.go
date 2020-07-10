package main

import (
	"database/sql"
	"flag"
	"fmt"
	config "github.com/go-ozzo/ozzo-config"
	_ "github.com/go-sql-driver/mysql"
	"github.com/tealeg/xlsx"
	"os"
	"runtime"
	"strconv"
	"time"
)

var columnSort = []string{
	"custCode",
	"custName",
	"businessId",
	"contractId",
	"certCode",
	"outAssignDate",
	"outAssignCode",
	"outAssignNum",
	"overdueTotalAmount",
	"installmentLoanAmount",
	"monthlyCapital",
	"monthlyInterest",
	"monthFee",
	"currentPenalty",
	"loanAmount",
	"paidAmount",
	"loanPeriod",
	"loanOutPeriod",
	"product",
	"repaymentAccountCode",
	"deliveryDate",
	"outAssignExpireDate",
	"overPeriod",
	"domicileTel",
	"unitTel",
	"mobile",
	"registerAddress",
	"residenceAddress",
	"unitAddress",
	"companyName",
	"immediateFamilyName",
	"immediateFamilyTel",
	"spouseName",
	"spouseTel",
	"contactsName",
	"contactsTel",
	"phoneNum1",
	"callTimes1",
	"phoneNum2",
	"callTimes2",
	"phoneNum3",
	"callTimes3",
	"phoneNum4",
	"callTimes4",
	"phoneNum5",
	"callTimes5",
	"phoneNum6",
	"callTimes6",
	"phoneNum7",
	"callTimes7",
	"phoneNum8",
	"callTimes8",
	"phoneNum9",
	"callTimes9",
	"phoneNum10",
	"callTimes10",
	"phoneNum11",
	"callTimes11",
	"phoneNum12",
	"callTimes12",
	"isApplyExecute",
	"salesCityName",
	"salesSubBranches",
	"subArea",
	"cityName",
	"divisionName",
	"currAssignCompany",
	"litigationStatus",
	"content",
	"collPeriod",
	"collContent",
}

var columnType = map[string]string {
	"custCode": "string",
	"custName": "string",
	"businessId": "string",
	"contractId": "string",
	"certCode": "string",
	"outAssignDate": "date",
	"outAssignCode": "string",
	"outAssignNum": "string",
	"overdueTotalAmount": "float",
	"installmentLoanAmount": "float",
	"monthlyCapital": "float",
	"monthlyInterest": "float",
	"monthFee": "float",
	"currentPenalty": "int",
	"loanAmount": "float",
	"paidAmount": "float",
	"loanPeriod": "int",
	"loanOutPeriod": "int",
	"product": "string",
	"repaymentAccountCode": "string",
	"deliveryDate": "date",
	"outAssignExpireDate": "date",
	"overPeriod": "int",
	"domicileTel": "string",
	"unitTel": "string",
	"mobile": "string",
	"registerAddress": "string",
	"residenceAddress": "string",
	"unitAddress": "string",
	"companyName": "string",
	"immediateFamilyName": "string",
	"immediateFamilyTel": "string",
	"spouseName": "string",
	"spouseTel": "string",
	"contactsName": "string",
	"contactsTel": "string",
	"phoneNum1": "string",
	"callTimes1": "int",
	"phoneNum2": "string",
	"callTimes2": "int",
	"phoneNum3": "string",
	"callTimes3": "int",
	"phoneNum4": "string",
	"callTimes4": "int",
	"phoneNum5": "string",
	"callTimes5": "int",
	"phoneNum6": "string",
	"callTimes6": "int",
	"phoneNum7": "string",
	"callTimes7": "int",
	"phoneNum8": "string",
	"callTimes8": "int",
	"phoneNum9": "string",
	"callTimes9": "int",
	"phoneNum10": "string",
	"callTimes10": "int",
	"phoneNum11": "string",
	"callTimes11": "int",
	"phoneNum12": "string",
	"callTimes12": "int",
	"isApplyExecute": "string",
	"salesCityName": "string",
	"salesSubBranches": "string",
	"subArea": "string",
	"cityName": "string",
	"divisionName": "string",
	"currAssignCompany": "string",
	"litigationStatus": "string",
	"content": "string",
	"collPeriod": "int",
	"collContent": "string",
}

func readData() []map[string]string {
	sqlText := `select * from test_data`
	db, err := sql.Open("mysql", dsn)
	if err != nil {
		panic(err)
	}
	defer db.Close()

	rows, err := db.Query(sqlText)
	if err != nil {
		panic(err)
	}
	defer rows.Close()

	columns, err := rows.Columns()
	if err != nil {
		panic(err)
	}
	values := make([]sql.RawBytes, len(columns))
	scanArgs := make([]interface{}, len(values))

	for i := range values {
		scanArgs[i] = &values[i]
	}
	data := make([]map[string]string, 0)
	for rows.Next() {
		err = rows.Scan(scanArgs...)
		if err != nil {
			panic(err.Error())
		}
		var value string
		valueMap := make(map[string]string, 0)
		for i, col := range values {
			if col == nil {
				value = ""
			} else {
				value = string(col)
			}
			valueMap[columns[i]] = value
		}
		data = append(data, valueMap)
	}
	return data
}

func export_process(data []map[string]string) {
	var file *xlsx.File
	file = xlsx.NewFile()
	sheet, err := file.AddSheet("test_data")
	if err != nil {
		fmt.Printf(err.Error())
	}

	//冻结
	sheet.SheetViews = []xlsx.SheetView{{Pane: &xlsx.Pane{
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
		State:       "frozen",
	}}}


	//处理表头
	xlsxRow := sheet.AddRow()
	for _, k := range columnSort {
		xlsxCell := xlsxRow.AddCell()
		xlsxCell.Value = k
		xlsxCell.GetStyle().Font.Bold = true
	}
	xlsxRow.SetHeight(20.0)

	for _, valueMap := range data {
		xlsxRow := sheet.AddRow()
		xlsxRow.SetHeight(20.0)

		for _, k := range columnSort {
			if _, ok := valueMap[k]; ok {
				xlsxCell := xlsxRow.AddCell()
				if columnType[k] == "int" {
					if valueMap[k] == "" {
						xlsxCell.SetString(valueMap[k])
						xlsxCell.SetFormat("general")
					} else {
						vInt64, err := strconv.ParseInt(valueMap[k], 10, 64)
						if err != nil {
							xlsxCell.SetString(valueMap[k])
							xlsxCell.SetFormat("general")
						} else {
							xlsxCell.SetInt64(vInt64)
							xlsxCell.SetFormat("general")
						}
					}
				} else if columnType[k] == "float" {
					if valueMap[k] == "" {
						xlsxCell.SetString(valueMap[k])
						xlsxCell.SetFormat("general")
					} else {
						vFloat64, err := strconv.ParseFloat(valueMap[k], 64)
						if err != nil {
							xlsxCell.SetString(valueMap[k])
							xlsxCell.SetFormat("general")
						} else {
							xlsxCell.SetFloat(vFloat64)
							xlsxCell.SetFormat("0.00")
						}
					}
				} else if columnType[k] == "date" {
					const Layout = "2006-01-02"//时间常量
					loc, _ := time.LoadLocation("Asia/Shanghai")
					t, err := time.ParseInLocation(Layout, valueMap[k], loc)
					if err != nil {
						xlsxCell.SetString(valueMap[k])
						xlsxCell.SetFormat("general")
					} else {
						xlsxCell.SetString(t.Format(Layout))
						xlsxCell.SetFormat("yyyy-mm-dd")
					}
				} else if columnType[k] == "string" {
					xlsxCell.SetString(valueMap[k])
					xlsxCell.SetFormat("general")
				}
			}
		}
	}

	err = file.Save("./golang.xlsx")
	if err != nil {
		panic(err)
	}
}

var dsn string

func main() {
	var confFilename *string = flag.String("conf", "", "specify a yaml conf filepath")
	flag.Parse()

	c := config.New()
	c.Load(*confFilename)
	if c.Get("database") == nil {
		fmt.Println("conf file not find database section or yaml format error")
		os.Exit(3)
	}

	maxproces := c.GetInt("maxproces", 0)

	dsn = c.GetString("database", "")

	if maxproces != 0 {
		runtime.GOMAXPROCS(maxproces)
	}

	data := readData()
	export_process(data)
}