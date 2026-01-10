package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	// 创建一个新的Excel文件
	f := excelize.NewFile()
	defer f.Close()

	// 设置Sheet名称
	sheetName := "报价单"
	f.SetSheetName("Sheet1", sheetName)

	// 设置列宽
	f.SetColWidth(sheetName, "A", "J", 15)
	f.SetColWidth(sheetName, "B", "B", 8)
	f.SetColWidth(sheetName, "D", "D", 20)

	// 标题样式
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 16,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})

	// 表头样式
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 12,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"#E0EBF5"},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})

	// 数据行样式
	dataStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical: "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})

	// 金额样式
	amountStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "right",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})

	// 公司名称
	f.SetCellValue(sheetName, "A1", "稻壳科技有限公司")
	f.SetCellStyle(sheetName, "A1", "J1", titleStyle)
	f.MergeCell(sheetName, "A1", "J1")

	// 报价单标题
	f.SetCellValue(sheetName, "A2", "报价单")
	f.SetCellStyle(sheetName, "A2", "J2", titleStyle)
	f.MergeCell(sheetName, "A2", "J2")

	// 公司信息
	f.SetCellValue(sheetName, "A3", "公司地址：XXXX稻壳科技有限公司")
	f.SetCellValue(sheetName, "A4", "报价说明：此为报价单说明，如有疑问，联系相关负责人！")
	f.SetCellValue(sheetName, "E3", "联系电话：0000-0000-0000")
	f.SetCellValue(sheetName, "G3", "地址：XXXX常州路15号")

	// 表头
	headers := []string{"品名", "产品图片", "规格", "材质说明", "颜色", "数量", "单价", "总价", "备注"}
	for i, header := range headers {
		cell := fmt.Sprintf("%c6", 'A'+i)
		f.SetCellValue(sheetName, cell, header)
		f.SetCellStyle(sheetName, cell, cell, headerStyle)
	}

	// 合并表头
	f.MergeCell(sheetName, "A5", "A6")
	f.MergeCell(sheetName, "B5", "B6")
	f.MergeCell(sheetName, "C5", "C6")
	f.MergeCell(sheetName, "D5", "D6")
	f.MergeCell(sheetName, "E5", "E6")
	f.MergeCell(sheetName, "F5", "F6")
	f.MergeCell(sheetName, "G5", "G6")
	f.MergeCell(sheetName, "H5", "H6")
	f.MergeCell(sheetName, "I5", "I6")

	// 示例数据
	exampleData := []struct {
		ProductName  string
		Specification string
		Material     string
		Color        string
		Quantity     int
		UnitPrice    float64
		TotalPrice   float64
		Remark       string
	}{
		{"大班台", "2400*2000*750", "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面", "黑色", 2, 10141.00, 20282.00, ""},
		{"文件柜", "2400*450*2000", "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面", "黑色", 3, 10716.00, 32148.00, ""},
		{"会客桌", "2000*800*750", "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面", "黑色", 6, 4500.00, 27000.00, ""},
		{"会客椅", "常规", "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面", "黑色", 1, 400.00, 400.00, ""},
		{"中式隔断", "定制2800*2100", "采用行列式手法，风格、功能设计", "不锈钢包边", 6, 1200.00, 7200.00, ""},
	}

	// 填充示例数据
	for i, data := range exampleData {
		row := i + 7
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), data.ProductName)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), data.Specification)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), data.Material)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), data.Color)
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), data.Quantity)
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), data.UnitPrice)
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), data.TotalPrice)
		f.SetCellValue(sheetName, fmt.Sprintf("I%d", row), data.Remark)

		// 设置样式
		f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("I%d", row), dataStyle)
		f.SetCellStyle(sheetName, fmt.Sprintf("G%d", row), fmt.Sprintf("H%d", row), amountStyle)
	}

	// 合计行
	totalRow := len(exampleData) + 7
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", totalRow), "合计（金额大写）：")
	f.SetCellValue(sheetName, fmt.Sprintf("E%d", totalRow), "柒仟贰佰元整")
	f.SetCellValue(sheetName, fmt.Sprintf("H%d", totalRow), "小计：")
	f.SetCellValue(sheetName, fmt.Sprintf("I%d", totalRow), 72000.00)

	// 优惠价行
	discountRow := totalRow + 1
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", discountRow), "此价格含运输，安装，增值税发票")
	f.SetCellValue(sheetName, fmt.Sprintf("H%d", discountRow), "优惠价：")
	f.SetCellValue(sheetName, fmt.Sprintf("I%d", discountRow), 50400.00)

	// 单位信息行
	companyRow := discountRow + 3
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", companyRow), "报价单位（盖章）：")
	f.SetCellValue(sheetName, fmt.Sprintf("F%d", companyRow), "询价单位（盖章）：")

	// 签名行
	signRow := companyRow + 1
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", signRow), "报价单位（签名）：")
	f.SetCellValue(sheetName, fmt.Sprintf("F%d", signRow), "询价单位（签名）：")

	// 设置合计行和优惠价行样式
	f.SetCellStyle(sheetName, fmt.Sprintf("A%d", totalRow), fmt.Sprintf("I%d", totalRow), dataStyle)
	f.SetCellStyle(sheetName, fmt.Sprintf("A%d", discountRow), fmt.Sprintf("I%d", discountRow), dataStyle)
	f.SetCellStyle(sheetName, fmt.Sprintf("G%d", totalRow), fmt.Sprintf("I%d", totalRow), amountStyle)
	f.SetCellStyle(sheetName, fmt.Sprintf("G%d", discountRow), fmt.Sprintf("I%d", discountRow), amountStyle)

	// 合并单元格
	f.MergeCell(sheetName, fmt.Sprintf("A%d", totalRow), fmt.Sprintf("D%d", totalRow))
	f.MergeCell(sheetName, fmt.Sprintf("E%d", totalRow), fmt.Sprintf("G%d", totalRow))
	f.MergeCell(sheetName, fmt.Sprintf("A%d", discountRow), fmt.Sprintf("G%d", discountRow))

	// 保存文件
	if err := f.SaveAs("templates/excel/quote.xlsx"); err != nil {
		fmt.Printf("保存文件失败: %v\n", err)
		return
	}

	fmt.Println("成功创建报价单模板文件: templates/excel/quote.xlsx")
}
