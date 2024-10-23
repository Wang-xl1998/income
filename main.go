package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// 清理收益字段，去掉"￥"符号等
func cleanRevenueValue(value string) (float64, error) {
	reg := regexp.MustCompile(`[^\d.]`)
	cleanedValue := reg.ReplaceAllString(value, "")
	if cleanedValue == "" {
		return 0, nil
	}
	return strconv.ParseFloat(cleanedValue, 64)
}

// 从文件名中提取年份和月份，格式 "2021区域结算单(03.01-03.31)"
func extractYearMonth(filePath string) (int, int, error) {
	base := filepath.Base(filePath)
	re := regexp.MustCompile(`(\d{4})区域结算单\((\d{2})`)
	matches := re.FindStringSubmatch(base)
	if len(matches) < 3 {
		return 0, 0, fmt.Errorf("无法从文件名提取年份和月份: %s", filePath)
	}
	year, _ := strconv.Atoi(matches[1])
	month, _ := strconv.Atoi(matches[2])
	return year, month, nil
}

// 处理收益每日明细sheet，获取站点总收益
func processDailySheet(filePath string) (map[string]float64, map[string]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, nil, err
	}
	defer f.Close()

	rows, err := f.GetRows("收益每日明细")
	if err != nil {
		return nil, nil, err
	}

	if len(rows) < 4 {
		return nil, nil, fmt.Errorf("文件行数不足")
	}

	stationNames := rows[2][1:]         // 第三行（跳过第一列）
	stationCodes := rows[3][1:]         // 第四行（跳过第一列）
	summaryRow := rows[len(rows)-1][1:] // 最后一行的总收益

	totalRevenue := make(map[string]float64)
	stationInfo := make(map[string]string)

	for i := 0; i < len(stationCodes); i++ {
		if i < len(summaryRow) {
			revenue, err := cleanRevenueValue(summaryRow[i])
			if err != nil {
				log.Printf("无法转换总收益: %v", err)
				continue
			}
			totalRevenue[stationCodes[i]] = revenue
			stationInfo[stationCodes[i]] = stationNames[i]
		}
	}

	return totalRevenue, stationInfo, nil
}

// 处理收益分项明细sheet，获取削峰填谷、效率提升和其他收益
func processSubItemSheet(filePath string) (map[string]map[string]float64, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows("收益分项明细")
	if err != nil {
		return nil, err
	}

	rows1, err := f.GetRows("收益每日明细")
	if err != nil {
		return nil, err
	}

	if len(rows) < 5 {
		return nil, fmt.Errorf("文件行数不足")
	}

	stationCodes := rows1[3][1:]
	peakValleyRow := rows[len(rows)-3][3:] // 削峰填谷收益
	efficiencyRow := rows[len(rows)-2][3:] // 效率提升收益
	otherRow := rows[len(rows)-1][3:]      // 其他收益

	subItems := make(map[string]map[string]float64)

	for i := 0; i < len(stationCodes); i++ {
		subItems[stationCodes[i]] = map[string]float64{}

		peakValley, err := cleanRevenueValue(peakValleyRow[i])
		if err == nil {
			subItems[stationCodes[i]]["削峰填谷收益"] = peakValley
		}

		efficiency, err := cleanRevenueValue(efficiencyRow[i])
		if err == nil {
			subItems[stationCodes[i]]["效率提升收益"] = efficiency
		}

		other, err := cleanRevenueValue(otherRow[i])
		if err == nil {
			subItems[stationCodes[i]]["其他收益"] = other
		}
	}

	return subItems, nil
}

func main() {
	dir := "./converted_files"
	// dir := "D:/结算单/利天万事202101-202409结算单"

	finalResults := make(map[string]map[int]float64)
	peakValleyResults := make(map[string]map[int]float64)
	efficiencyResults := make(map[string]map[int]float64)
	otherResults := make(map[string]map[int]float64)
	finalStationNames := make(map[string]string)
	monthlyResults := make(map[string]map[int]map[int]map[string]float64)

	err := filepath.WalkDir(dir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if filepath.Ext(path) == ".xlsx" {
			year, month, err := extractYearMonth(path)
			if err != nil {
				log.Printf("无法提取年份和月份: %v\n", err)
				return nil
			}

			// 获取收益每日明细的总收益
			totalRevenue, stationInfo, err := processDailySheet(path)
			if err != nil {
				log.Printf("处理文件 %s 时出错: %v\n", path, err)
				return nil
			}

			// 获取收益分项明细的削峰填谷、效率提升、其他收益
			subItemRevenue, err := processSubItemSheet(path)
			if err != nil {
				log.Printf("处理文件 %s 时出错: %v\n", path, err)
				return nil
			}

			for code, revenue := range totalRevenue {
				if finalResults[code] == nil {
					finalResults[code] = make(map[int]float64)
					peakValleyResults[code] = make(map[int]float64)
					efficiencyResults[code] = make(map[int]float64)
					otherResults[code] = make(map[int]float64)
				}
				finalResults[code][year] += revenue
				peakValleyResults[code][year] += subItemRevenue[code]["削峰填谷收益"]
				efficiencyResults[code][year] += subItemRevenue[code]["效率提升收益"]
				otherResults[code][year] += subItemRevenue[code]["其他收益"]
				finalStationNames[code] = stationInfo[code]

				// 处理月度明细
				if monthlyResults[code] == nil {
					monthlyResults[code] = make(map[int]map[int]map[string]float64)
				}
				if monthlyResults[code][year] == nil {
					monthlyResults[code][year] = make(map[int]map[string]float64)
				}
				monthlyResults[code][year][month] = map[string]float64{
					"总收益":    revenue,
					"削峰填谷收益": subItemRevenue[code]["削峰填谷收益"],
					"效率提升收益": subItemRevenue[code]["效率提升收益"],
					"其他收益":   subItemRevenue[code]["其他收益"],
				}
			}
		}
		return nil
	})

	if err != nil {
		log.Fatal(err)
	}

	// 对站点编码进行排序
	sortedCodes := make([]string, 0, len(finalResults))
	for code := range finalResults {
		sortedCodes = append(sortedCodes, code)
	}
	sort.Strings(sortedCodes)

	// 创建新的Excel文件
	newFile := excelize.NewFile()
	summarySheet := "站点汇总"
	detailSheet := "年度明细"
	monthlyDetailSheet := "月度明细"
	newFile.NewSheet(summarySheet)
	newFile.NewSheet(detailSheet)
	newFile.NewSheet(monthlyDetailSheet)

	// 设置汇总表的表头
	// 设置汇总表的表头 (接着继续前面的代码)
	newFile.SetCellValue(summarySheet, "A1", "站点名称")
	newFile.SetCellValue(summarySheet, "B1", "站点编码")
	newFile.SetCellValue(summarySheet, "C1", "总收益")
	newFile.SetCellValue(summarySheet, "D1", "削峰填谷收益")
	newFile.SetCellValue(summarySheet, "E1", "效率提升收益")
	newFile.SetCellValue(summarySheet, "F1", "其它收益")

	// 填充汇总表
	rowIndex := 2
	for _, code := range sortedCodes {
		name := finalStationNames[code]
		totalRevenue := 0.0
		peakValleyTotal := 0.0
		efficiencyTotal := 0.0
		otherTotal := 0.0

		// 遍历每年数据进行累加
		for year := range finalResults[code] {
			totalRevenue += finalResults[code][year]
			peakValleyTotal += peakValleyResults[code][year]
			efficiencyTotal += efficiencyResults[code][year]
			otherTotal += otherResults[code][year]
		}

		newFile.SetCellValue(summarySheet, fmt.Sprintf("A%d", rowIndex), name)
		newFile.SetCellValue(summarySheet, fmt.Sprintf("B%d", rowIndex), code)
		newFile.SetCellValue(summarySheet, fmt.Sprintf("C%d", rowIndex), totalRevenue)
		newFile.SetCellValue(summarySheet, fmt.Sprintf("D%d", rowIndex), peakValleyTotal)
		newFile.SetCellValue(summarySheet, fmt.Sprintf("E%d", rowIndex), efficiencyTotal)
		newFile.SetCellValue(summarySheet, fmt.Sprintf("F%d", rowIndex), otherTotal)
		rowIndex++
	}

	// 设置年度明细表的表头
	newFile.SetCellValue(detailSheet, "A1", "站点名称")
	newFile.SetCellValue(detailSheet, "B1", "站点编码")

	// 获取所有年份并排序
	years := map[int]bool{}
	for _, yearlyData := range finalResults {
		for year := range yearlyData {
			years[year] = true
		}
	}
	yearList := []int{}
	for year := range years {
		yearList = append(yearList, year)
	}
	sort.Ints(yearList)

	// 设置年度明细表的表头，按年份添加列
	colIndex := 3 // 从第三列开始
	for _, year := range yearList {
		// 生成列号
		baseCol, _ := excelize.ColumnNumberToName(colIndex)
		newFile.SetCellValue(detailSheet, fmt.Sprintf("%s1", baseCol), fmt.Sprintf("%d年总收益", year))

		peakValleyCol, _ := excelize.ColumnNumberToName(colIndex + 1)
		newFile.SetCellValue(detailSheet, fmt.Sprintf("%s1", peakValleyCol), fmt.Sprintf("%d年削峰填谷收益", year))

		efficiencyCol, _ := excelize.ColumnNumberToName(colIndex + 2)
		newFile.SetCellValue(detailSheet, fmt.Sprintf("%s1", efficiencyCol), fmt.Sprintf("%d年效率提升收益", year))

		otherCol, _ := excelize.ColumnNumberToName(colIndex + 3)
		newFile.SetCellValue(detailSheet, fmt.Sprintf("%s1", otherCol), fmt.Sprintf("%d年其它收益", year))

		colIndex += 4 // 每次增加4列
	}

	// 填充年度明细表
	rowIndex = 2
	for _, code := range sortedCodes {
		name := finalStationNames[code]
		newFile.SetCellValue(detailSheet, fmt.Sprintf("A%d", rowIndex), name)
		newFile.SetCellValue(detailSheet, fmt.Sprintf("B%d", rowIndex), code)

		colIndex = 3
		for _, year := range yearList {
			// 生成列号
			baseCol, _ := excelize.ColumnNumberToName(colIndex)
			newFile.SetCellValue(detailSheet, fmt.Sprintf("%s%d", baseCol, rowIndex), finalResults[code][year])

			peakValleyCol, _ := excelize.ColumnNumberToName(colIndex + 1)
			newFile.SetCellValue(detailSheet, fmt.Sprintf("%s%d", peakValleyCol, rowIndex), peakValleyResults[code][year])

			efficiencyCol, _ := excelize.ColumnNumberToName(colIndex + 2)
			newFile.SetCellValue(detailSheet, fmt.Sprintf("%s%d", efficiencyCol, rowIndex), efficiencyResults[code][year])

			otherCol, _ := excelize.ColumnNumberToName(colIndex + 3)
			newFile.SetCellValue(detailSheet, fmt.Sprintf("%s%d", otherCol, rowIndex), otherResults[code][year])

			colIndex += 4
		}
		rowIndex++
	}

	// 设置月度明细表的表头
	newFile.SetCellValue(monthlyDetailSheet, "A1", "站点名称")
	newFile.SetCellValue(monthlyDetailSheet, "B1", "站点编码")

	// 获取所有年份和月份并排序
	yearsMonths := map[int]map[int]bool{}
	for _, yearData := range monthlyResults {
		for year, monthData := range yearData {
			if yearsMonths[year] == nil {
				yearsMonths[year] = map[int]bool{}
			}
			for month := range monthData {
				yearsMonths[year][month] = true
			}
		}
	}

	yearList = []int{}
	for year := range yearsMonths {
		yearList = append(yearList, year)
	}
	sort.Ints(yearList)

	// 设置月度明细表的表头，按“年份 + 月份”添加列
	colIndex = 3 // 从Excel第三列（C列）开始
	for _, year := range yearList {
		months := []int{}
		for month := range yearsMonths[year] {
			months = append(months, month)
		}
		sort.Ints(months)

		for _, month := range months {
			monthLabel := fmt.Sprintf("%d年%d月", year, month)

			// 使用 excelize.ColumnNumberToName 来转换列号
			baseCol, _ := excelize.ColumnNumberToName(colIndex)
			newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s1", baseCol), fmt.Sprintf("%s总收益", monthLabel))

			peakValleyCol, _ := excelize.ColumnNumberToName(colIndex + 1)
			newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s1", peakValleyCol), fmt.Sprintf("%s削峰填谷收益", monthLabel))

			efficiencyCol, _ := excelize.ColumnNumberToName(colIndex + 2)
			newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s1", efficiencyCol), fmt.Sprintf("%s效率提升收益", monthLabel))

			otherCol, _ := excelize.ColumnNumberToName(colIndex + 3)
			newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s1", otherCol), fmt.Sprintf("%s其它收益", monthLabel))

			colIndex += 4 // 每次增加4列
		}
	}

	// 填充月度明细表
	rowIndex = 2
	for _, code := range sortedCodes {
		name := finalStationNames[code]
		newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("A%d", rowIndex), name)
		newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("B%d", rowIndex), code)

		colIndex = 3
		for _, year := range yearList {
			months := []int{}
			for month := range yearsMonths[year] {
				months = append(months, month)
			}
			sort.Ints(months)

			for _, month := range months {
				revenues := monthlyResults[code][year][month]

				// 检查 revenues 是否为空
				if revenues == nil {
					colIndex += 4
					continue
				}

				baseCol, _ := excelize.ColumnNumberToName(colIndex)
				newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s%d", baseCol, rowIndex), revenues["总收益"])

				peakValleyCol, _ := excelize.ColumnNumberToName(colIndex + 1)
				newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s%d", peakValleyCol, rowIndex), revenues["削峰填谷收益"])

				efficiencyCol, _ := excelize.ColumnNumberToName(colIndex + 2)
				newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s%d", efficiencyCol, rowIndex), revenues["效率提升收益"])

				otherCol, _ := excelize.ColumnNumberToName(colIndex + 3)
				newFile.SetCellValue(monthlyDetailSheet, fmt.Sprintf("%s%d", otherCol, rowIndex), revenues["其他收益"])

				colIndex += 4
			}
		}
		rowIndex++
	}

	// 保存 Excel 文件
	if err := newFile.SaveAs("站点年度和月度收益明细.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("站点年度和月度收益明细已成功保存至 '站点年度和月度收益明细.xlsx'")
}
