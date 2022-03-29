package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

const (
	FILENAME          = "AIMdex score data.xlsx"
	KEEP_GOING_PROMPT = "\n要继续计算吗? [若要继续，将新的文件改名为AIMdex score data.xlsx，放入与上一个文件同一个路径后在这里输入yes并按回车键]\n"
	COLOR_BLUE        = "\033[1;34m%s\033[0m"
	COLOR_LIGHT_BLUE  = "\033[0;36m%s\033[0m"
	COLOR_RED         = "\033[1;31m%s\033[0m"
	COLOR_PURPLE      = "\033[4;35m%s\033[0m"
)

func main() {
	keepGoing := true
	for keepGoing {
		err := run()
		if err != nil {
			log.Fatal(err)
		}
		keepGoing = (promptUserToInput(KEEP_GOING_PROMPT) == "yes")
	}

	fmt.Print("运行结束。\n")
}

func run() error {
	path, err := os.Getwd()
	if err != nil {
		log.Println(err)
	}

	fmt.Print("欢迎使用AIMdex计算器。\n尝试读取文件：", FILENAME, "\n请保证此xlsx正在路径：", path, "里\n")

	file, err := xlsx.OpenFile(FILENAME)
	if err != nil {
		fmt.Print("读取文件错误，或者文件不存在，请检查" + FILENAME)
		return err
	}

	if len(file.Sheet) < 2 {
		return fmt.Errorf("这个文件不足2个sheet。\n")
	}

	for _, row := range file.Sheets[0].Rows {
		if len(row.Cells) < 10 {
			continue
		}

		d1 := row.Cells[7].String()
		d2 := row.Cells[8].String()
		d3 := row.Cells[9].String()

		data1, err := strconv.Atoi(d1)
		if err != nil {
			continue
		}

		data2, err := strconv.Atoi(d2)
		if err != nil {
			continue
		}

		data3, err := strconv.Atoi(d3)
		if err != nil {
			continue
		}

		//fmt.Print("\n", i, ": ", data1, " ", data2, " ", data3, "\n")

		for _, row2 := range file.Sheets[1].Rows {
			if1 := row2.Cells[0].String()
			if2 := row2.Cells[1].String()
			if3 := row2.Cells[2].String()
			score := row2.Cells[3].String()

			// 如果符合3个条件，score填写到sheet0里
			if parseCompareNum(data1, if1) && parseCompareNum(data2, if2) && parseCompareNum(data3, if3) {
				//fmt.Print(j, ": ", if1, " ", if2, " ", if3, " ", score, "\n")
				for len(row.Cells) <= 13 {
					row.AddCell()
				}
				row.Cells[13].SetString(score)
				//fmt.Print("writing to cell N with score ", score, "\n")
				break
			}
		}
	}

	err = file.Save(FILENAME)
	fmt.Print("\n成功读取与书写。分数已写入第一个表格的N栏。\n")
	return err
}

func promptUserToInput(request string) string {
	// When user execute the binary, prompt them to input a file path to read the photos from
	fmt.Printf(COLOR_BLUE, request)
	var answer string
	fmt.Scanln(&answer)
	return answer
}

// num = 5, s = ">=10"
func parseCompareNum(num int, s string) bool {
	if strings.Contains(s, ">=") {
		nString := strings.Split(s, ">=")[1]
		n, err := strconv.Atoi(nString)
		if err != nil {
			fmt.Print("处理 " + s + " 时出现错误")
		}

		return num >= n
	}

	if strings.Contains(s, ">") {
		nString := strings.Split(s, ">")[1]
		n, err := strconv.Atoi(nString)
		if err != nil {
			fmt.Print("处理 " + s + " 时出现错误")
		}

		return num > n
	}

	if strings.Contains(s, "=") {
		nString := strings.Split(s, "=")[1]
		n, err := strconv.Atoi(nString)
		if err != nil {
			fmt.Print("处理 " + s + " 时出现错误")
		}

		return num == n
	}

	if strings.Contains(s, "<") {
		nString := strings.Split(s, "<")[1]
		n, err := strconv.Atoi(nString)
		if err != nil {
			fmt.Print("处理 " + s + " 时出现错误")
		}

		return num < n
	}

	if strings.Contains(s, "<=") {
		nString := strings.Split(s, "<=")[1]
		n, err := strconv.Atoi(nString)
		if err != nil {
			fmt.Print("处理 " + s + " 时出现错误")
		}

		return num <= n
	}

	return false
}
