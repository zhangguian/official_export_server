package export

import (
	"fmt"
	"strings"

	"office-export-server/internal/model"
	"office-export-server/internal/service/template"

	"github.com/xuri/excelize/v2"
)

// ExcelService Excel导出服务
type ExcelService struct {
	templateService template.TemplateService
}

// NewExcelService 创建Excel导出服务实例
func NewExcelService(templateService template.TemplateService) *ExcelService {
	return &ExcelService{
		templateService: templateService,
	}
}

// ExportExcel 导出Excel文件
func (s *ExcelService) ExportExcel(req *model.ExportRequest) ([]byte, error) {
	// 获取模板ID（直接从请求根级别获取，而不是从Data字段）
	templateID := req.TemplateID
	if templateID == "" {
		templateID = "default"
	}

	// 获取模板路径
	templatePath, err := s.templateService.GetTemplatePath(templateID, "excel")
	if err != nil {
		return nil, fmt.Errorf("failed to get template path: %v", err)
	}

	// 打开模板文件
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open template file: %v", err)
	}
	defer f.Close()

	// 统一处理sheets数组，前端必须传递sheets:[]结构
	var sheets []interface{}
	var ok bool
	if sheets, ok = req.Data["sheets"].([]interface{}); !ok || len(sheets) == 0 {
		return nil, fmt.Errorf("前端必须传递sheets:[]数组结构，且数组不能为空")
	}

	// 遍历sheets数组，为每个sheet创建新的sheet页
	defaultSheet := f.GetSheetName(0)
	for i, sheetData := range sheets {
		sheetMap, ok := sheetData.(map[string]interface{})
		if !ok {
			continue
		}

		// 获取sheet名称
		sheetName, ok := sheetMap["name"].(string)
		if !ok || sheetName == "" {
			sheetName = fmt.Sprintf("Sheet%d", i+1)
		}

		// 创建新的sheet页
		_, err := f.NewSheet(sheetName)
		if err != nil {
			return nil, fmt.Errorf("failed to create new sheet: %v", err)
		}

		// 填充当前sheet的数据
		// 创建临时请求对象，包含当前sheet的数据
		tempReq := &model.ExportRequest{
			TemplateID: templateID,
			DataType:   req.DataType,
			Data:       sheetMap,
		}

		// 填充数据到当前sheet
		if err := s.fillTemplateData(f, sheetName, templateID, tempReq); err != nil {
			return nil, fmt.Errorf("failed to fill template data for sheet %s: %v", sheetName, err)
		}
	}

	// 删除模板中的默认sheet（如果创建了新的sheet）
	if len(sheets) > 0 && f.GetSheetName(0) == defaultSheet {
		if err := f.DeleteSheet(defaultSheet); err != nil {
			return nil, fmt.Errorf("failed to delete default sheet: %v", err)
		}
	}

	// 保存为二进制数据
	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("failed to write excel to buffer: %v", err)
	}
	return buf.Bytes(), nil
}

// fillTemplateData 根据模板类型填充数据
func (s *ExcelService) fillTemplateData(f *excelize.File, sheetName, templateID string, req *model.ExportRequest) error {
	// 根据模板ID选择不同的数据填充逻辑
	switch templateID {
	case "budget":
		return s.fillBudgetTemplateData(f, sheetName, req)
	case "simple":
		return s.fillSimpleTemplateData(f, sheetName, req)
	case "quote":
		return s.fillQuoteTemplateData(f, sheetName, req)
	default:
		return s.fillDefaultTemplateData(f, sheetName, req)
	}
}

// fillDefaultTemplateData 填充默认模板数据
func (s *ExcelService) fillDefaultTemplateData(f *excelize.File, sheetName string, req *model.ExportRequest) error {
	// 处理预算汇总表模板
	return s.processBudgetTemplate(f, sheetName, req)
}

// fillBudgetTemplateData 填充预算汇总表模板数据
func (s *ExcelService) fillBudgetTemplateData(f *excelize.File, sheetName string, req *model.ExportRequest) error {
	// 与默认模板相同，仅作为示例
	return s.fillDefaultTemplateData(f, sheetName, req)
}

// fillSimpleTemplateData 填充简单模板数据
func (s *ExcelService) fillSimpleTemplateData(f *excelize.File, sheetName string, req *model.ExportRequest) error {
	// 设置列宽
	f.SetColWidth(sheetName, "A", "D", 20)

	// 添加标题
	f.SetCellValue(sheetName, "A1", "简单报表")
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 16,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	f.MergeCell(sheetName, "A1", "D1")
	f.SetCellStyle(sheetName, "A1", "D1", titleStyle)

	// 添加表头
	headers := []string{"序号", "名称", "数量", "金额"}
	for i, header := range headers {
		cell := fmt.Sprintf("%c2", 'A'+i)
		f.SetCellValue(sheetName, cell, header)
	}

	// 表头样式
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
	})
	f.SetCellStyle(sheetName, "A2", "D2", headerStyle)

	// 数据样式
	dataStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical: "center",
			WrapText: true,
		},
	})

	// 添加数据行
	items, ok := req.Data["items"].([]interface{})
	if ok {
		for i, item := range items {
			row := i + 3
			itemMap, _ := item.(map[string]interface{})
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), i+1)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), itemMap["品牌"])
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), itemMap["工程量"])
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), itemMap["预算价"])
			// 设置数据样式
			f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row), dataStyle)
		}
	}

	// 动态调整数据行高，适应内容
	// if items != nil && len(items) > 0 {
	// 	for i := 3; i <= len(items)+2; i++ {
	// 		// 对于每一行，根据内容长度动态调整行高
	// 		// 计算所需的行高（基础高度 + 内容长度/每行字符数 * 行高）
	// 		baseHeight := 20.0
	// 		// 获取该行的主要内容单元格（例如B列和D列）
	// 		contentCell := fmt.Sprintf("B%d", i)
	// 		content, _ := f.GetCellValue(sheetName, contentCell)
	// 		// 计算所需行数（假设每行能显示30个字符）
	// 		lines := len(content) / 30
	// 		if len(content)%30 > 0 {
	// 			lines++
	// 		}
	// 		// 设置行高（最小20，最大100）
	// 		height := baseHeight * float64(lines)
	// 		if height < 20 {
	// 			height = 20
	// 		} else if height > 100 {
	// 			height = 100
	// 		}
	// 		f.SetRowHeight(sheetName, i, height)
	// 	}
	// }

	return nil
}

// fillQuoteTemplateData 填充报价单模板数据
func (s *ExcelService) fillQuoteTemplateData(f *excelize.File, sheetName string, req *model.ExportRequest) error {
	// 数据行样式 - 添加自动换行
	dataStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical: "center",
			WrapText: true, // 自动换行
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
			WrapText:   true, // 自动换行
		},
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})

	// 获取数据
	items, ok := req.Data["items"].([]interface{})
	if !ok {
		// 如果没有数据，使用默认数据
		items = []interface{}{
			map[string]interface{}{
				"品名":   "大班台",
				"规格":   "2400*2000*750",
				"材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
				"颜色":   "黑色",
				"数量":   2,
				"单价":   10141.00,
				"总价":   20282.00,
				"备注":   "",
			},
			map[string]interface{}{
				"品名":   "文件柜",
				"规格":   "2400*450*2000",
				"材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
				"颜色":   "黑色",
				"数量":   3,
				"单价":   10716.00,
				"总价":   32148.00,
				"备注":   "",
			},
			map[string]interface{}{
				"品名":   "会客桌",
				"规格":   "2000*800*750",
				"材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
				"颜色":   "黑色",
				"数量":   6,
				"单价":   4500.00,
				"总价":   27000.00,
				"备注":   "",
			},
			map[string]interface{}{
				"品名":   "会客椅",
				"规格":   "常规",
				"材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
				"颜色":   "黑色",
				"数量":   1,
				"单价":   400.00,
				"总价":   400.00,
				"备注":   "",
			},
			map[string]interface{}{
				"品名":   "中式隔断",
				"规格":   "定制2800*2100",
				"材质说明": "采用行列式手法，风格、功能设计",
				"颜色":   "不锈钢包边",
				"数量":   6,
				"单价":   1200.00,
				"总价":   7200.00,
				"备注":   "",
			},
		}
	}

	// 填充数据行
	total := 0.0
	for i, item := range items {
		row := i + 7
		itemMap, _ := item.(map[string]interface{})

		// 填充数据
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), itemMap["品名"])
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), itemMap["规格"])
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), itemMap["材质说明"])
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), itemMap["颜色"])
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), itemMap["数量"])
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), itemMap["单价"])
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), itemMap["总价"])
		f.SetCellValue(sheetName, fmt.Sprintf("I%d", row), itemMap["备注"])

		// 累加总价
		if totalPrice, ok := itemMap["总价"].(float64); ok {
			total += totalPrice
		}

		// 设置样式
		f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("I%d", row), dataStyle)
		f.SetCellStyle(sheetName, fmt.Sprintf("G%d", row), fmt.Sprintf("H%d", row), amountStyle)
	}

	// 合计行
	totalRow := len(items) + 7
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", totalRow), "合计（金额大写）：")
	f.SetCellValue(sheetName, fmt.Sprintf("H%d", totalRow), "小计：")
	f.SetCellValue(sheetName, fmt.Sprintf("I%d", totalRow), total)

	// 优惠价行
	discountRow := totalRow + 1
	// 默认优惠价为原价的70%
	discountPrice := total * 0.7
	f.SetCellValue(sheetName, fmt.Sprintf("H%d", discountRow), "优惠价：")
	f.SetCellValue(sheetName, fmt.Sprintf("I%d", discountRow), discountPrice)

	// 设置合计行和优惠价行样式
	f.SetCellStyle(sheetName, fmt.Sprintf("G%d", totalRow), fmt.Sprintf("I%d", totalRow), amountStyle)
	f.SetCellStyle(sheetName, fmt.Sprintf("G%d", discountRow), fmt.Sprintf("I%d", discountRow), amountStyle)

	return nil
}

// processHeaders 处理表头
func (s *ExcelService) processHeaders(f *excelize.File, sheetName string, headers []interface{}) error {
	for rowIdx, headerRow := range headers {
		headerCells, ok := headerRow.([]interface{})
		if !ok {
			return fmt.Errorf("invalid header row format")
		}

		for colIdx, cellData := range headerCells {
			colStr, err := excelize.ColumnNumberToName(colIdx + 1)
			if err != nil {
				return err
			}

			cellValue, ok := cellData.(string)
			if !ok {
				cellValue = ""
			}

			cell := fmt.Sprintf("%s%d", colStr, rowIdx+1)
			f.SetCellValue(sheetName, cell, cellValue)

			// 设置表头样式
			style, err := f.NewStyle(&excelize.Style{
				Font: &excelize.Font{
					Bold: true,
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
			if err != nil {
				return err
			}
			f.SetCellStyle(sheetName, cell, cell, style)
		}
	}

	return nil
}

// processRows 处理数据行
func (s *ExcelService) processRows(f *excelize.File, sheetName string, rows []interface{}, headerRows int) error {
	for rowIdx, rowData := range rows {
		cells, ok := rowData.([]interface{})
		if !ok {
			return fmt.Errorf("invalid row data format")
		}

		for colIdx, cellData := range cells {
			colStr, err := excelize.ColumnNumberToName(colIdx + 1)
			if err != nil {
				return err
			}

			cell := fmt.Sprintf("%s%d", colStr, rowIdx+headerRows+1)
			f.SetCellValue(sheetName, cell, cellData)

			// 设置单元格样式
			style, err := f.NewStyle(&excelize.Style{
				Alignment: &excelize.Alignment{
					Vertical: "center",
					WrapText: true,
				},
				Border: []excelize.Border{
					{Type: "left", Color: "#000000", Style: 1},
					{Type: "top", Color: "#000000", Style: 1},
					{Type: "right", Color: "#000000", Style: 1},
					{Type: "bottom", Color: "#000000", Style: 1},
				},
			})
			if err != nil {
				return err
			}
			f.SetCellStyle(sheetName, cell, cell, style)
		}
	}

	return nil
}

// processMerges 处理合并单元格
func (s *ExcelService) processMerges(f *excelize.File, sheetName string, merges []interface{}) error {
	for _, mergeItem := range merges {
		mergeMap, ok := mergeItem.(map[string]interface{})
		if !ok {
			continue
		}

		startRow, _ := mergeMap["start_row"].(float64)
		startCol, _ := mergeMap["start_col"].(float64)
		endRow, _ := mergeMap["end_row"].(float64)
		endCol, _ := mergeMap["end_col"].(float64)

		startColStr, err := excelize.ColumnNumberToName(int(startCol) + 1)
		if err != nil {
			return err
		}

		endColStr, err := excelize.ColumnNumberToName(int(endCol) + 1)
		if err != nil {
			return err
		}

		// 使用MergeCell方法，需要指定起始和结束单元格
		startCell := fmt.Sprintf("%s%d", startColStr, int(startRow)+1)
		endCell := fmt.Sprintf("%s%d", endColStr, int(endRow)+1)
		if err := f.MergeCell(sheetName, startCell, endCell); err != nil {
			return fmt.Errorf("failed to merge cells: %v", err)
		}
	}

	return nil
}

// processBudgetTemplate 处理预算汇总表模板
func (s *ExcelService) processBudgetTemplate(f *excelize.File, sheetName string, req *model.ExportRequest) error {
	// 设置默认列宽
	columnWidths := map[string]float64{
		"A": 5,
		"B": 15,
		"C": 15,
		"D": 80,
		"E": 10,
		"F": 10,
		"G": 15,
		"H": 15,
	}
	for col, width := range columnWidths {
		f.SetColWidth(sheetName, col, col, width)
	}

	// 设置行高
	f.SetRowHeight(sheetName, 1, 30) // 标题行固定行高
	f.SetRowHeight(sheetName, 2, 25) // 项目信息行固定行高
	f.SetRowHeight(sheetName, 3, 25) // 客户经理行固定行高
	f.SetRowHeight(sheetName, 4, 25) // 表头行固定行高
	f.SetRowHeight(sheetName, 5, 25) // 表头行固定行高

	// 标题样式
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:  true,
			Size:  16,
			Color: "#000000",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})

	// 表头样式
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:  true,
			Size:  12,
			Color: "#000000",
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
			WrapText: true,
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
			WrapText:   true,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})

	// 合并单元格
	merges := []string{
		"A1:H1", // 主标题
		"A2:C2", // 项目名称
		"D2:G2", // 设计师
		"H2:H3", // 参考户型图
		"A3:C3", // 客户经理
		"D3:G3", // 设计师
		"A4:A5", // 序号
		"B4:B5", // 品牌
		"C4:C5", // 区域
		"D4:D5", // 系统说明
		"E4:E5", // 单位
		"F4:F5", // 工程量
		"G4:G5", // 预算价
		"H4:H5", // 单项预算合价
	}
	for _, merge := range merges {
		// 解析合并范围，拆分为起始和结束单元格
		parts := strings.Split(merge, ":")
		if len(parts) == 2 {
			f.MergeCell(sheetName, parts[0], parts[1])
		}
	}

	// 填写标题信息
	f.SetCellValue(sheetName, "A1", "全屋智能家居方案A套餐预算汇总表")
	f.SetCellStyle(sheetName, "A1", "H1", titleStyle)

	// 填写项目基本信息
	f.SetCellValue(sheetName, "A2", "项目名称：")
	f.SetCellValue(sheetName, "B2", "三房二厅全屋智能家居项目")
	f.SetCellValue(sheetName, "D2", "设计师：")
	f.SetCellValue(sheetName, "H2", "参考户型图")
	f.SetCellValue(sheetName, "A3", "客户经理：")
	f.SetCellValue(sheetName, "D3", "设计师：")

	// 插入参考户型图
	// 图片URL
	// imageURL := "https://www.zhijiayi.com/upload/houseConfig/20201228/3537592706212974.png"
	// // 从远程URL下载图片
	// resp, err := http.Get(imageURL)
	// if err != nil {
	// 	fmt.Printf("Failed to download reference house image: %v\n", err)
	// } else {
	// 	defer resp.Body.Close()
	// 	// 读取图片内容
	// 	imageBytes, err := ioutil.ReadAll(resp.Body)
	// 	if err != nil {
	// 		fmt.Printf("Failed to read reference house image: %v\n", err)
	// 	} else {
	// 		// 创建临时文件保存图片
	// 		tempFile, err := os.CreateTemp("", "house-image-*.png")
	// 		if err != nil {
	// 			fmt.Printf("Failed to create temp file: %v\n", err)
	// 		} else {
	// 			defer func() {
	// 				tempFile.Close()
	// 				os.Remove(tempFile.Name())
	// 			}()
	// 			// 写入图片内容
	// 			if _, err := tempFile.Write(imageBytes); err != nil {
	// 				fmt.Printf("Failed to write temp file: %v\n", err)
	// 			} else {
	// 				// 刷新文件缓冲区
	// 				tempFile.Sync()
	// 				// 设置图片位置和尺寸（H2:H3合并单元格）
	// 				printObject := true
	// 				// 设置行高，为图片留出空间
	// 				f.SetRowHeight(sheetName, 2, 120) // 设置第2行高度为120磅
	// 				f.SetRowHeight(sheetName, 3, 120) // 设置第3行高度为120磅

	// 				// 设置列宽，为图片留出空间
	// 				f.SetColWidth(sheetName, "H", "H", 15) // 设置H列宽度为15列宽

	// 				// 使用AddPicture方法，传入临时文件路径，调整缩放比例以达到预期宽高
	// 				if err := f.AddPicture(
	// 					sheetName,
	// 					"H2",
	// 					tempFile.Name(), // 临时文件路径
	// 					&excelize.GraphicOptions{
	// 						// 调整缩放比例，使图片达到预期大小
	// 						ScaleX:          0.6, // 水平缩放比例
	// 						ScaleY:          0.8, // 垂直缩放比例
	// 						OffsetX:         5,   // 轻微偏移，使图片居中
	// 						OffsetY:         5,   // 轻微偏移，使图片居中
	// 						PrintObject:     &printObject,
	// 						LockAspectRatio: false, // 不锁定宽高比，允许调整缩放比例
	// 					},
	// 				); err != nil {
	// 					// 如果图片插入失败，继续执行，不影响其他功能
	// 					fmt.Printf("Failed to add reference house image: %v\n", err)
	// 				}
	// 			}
	// 		}
	// 	}
	// }

	// 填写表头
	f.SetCellValue(sheetName, "A4", "序号")
	f.SetCellValue(sheetName, "B4", "品牌")
	f.SetCellValue(sheetName, "C4", "区域")
	f.SetCellValue(sheetName, "D4", "系统说明")
	f.SetCellValue(sheetName, "E4", "单位")
	f.SetCellValue(sheetName, "F4", "工程量")
	f.SetCellValue(sheetName, "G4", "预算价（元）")
	f.SetCellValue(sheetName, "H4", "单项预算合价（元）")

	// 设置表头样式
	f.SetCellStyle(sheetName, "A4", "H5", headerStyle)

	// 获取数据
	items, ok := req.Data["items"].([]interface{})
	if !ok {
		// 如果没有数据，使用示例数据
		items = []interface{}{
			map[string]interface{}{
				"序号":     1,
				"品牌":     "小米",
				"区域":     "全屋智能主控系统",
				"系统说明":   "1、AI智能语音、自定义设备各种场景（回家、离家、会客、就餐、休闲、阅读等模式），完美实现智能化体验。 2、智能品类包括：智能灯光、智能遮阳、智能空调，智能安防等； 3、最大优势及亮点 \"无缝接入米家APP、AI智能语音控制、轻成本、轻设计、轻方案、轻对接、轻落地、轻维护\"； 4、可以根据所需的智能开关与空调语音小助手进行DIY定制。",
				"单位":     "项",
				"工程量":    1,
				"预算价":    1928.00,
				"单项预算合价": 1928.00,
			},
			map[string]interface{}{
				"序号":     2,
				"品牌":     "FSXRT",
				"区域":     "智能灯光",
				"系统说明":   "1、自定义色温：智能双色温（2700~6500K的灯具，可以根据需求DIY自定义设置色温参数； 2、控制方式：单灯控制、回路控制、互控、集成控制、远程控制等； 3、自定义氛围场景：娱乐、聚会、休闲、会客等灯光场景。",
				"单位":     "项",
				"工程量":    1,
				"预算价":    6759.00,
				"单项预算合价": 6759.00,
			},
		}
	}

	// 填写数据行
	startRow := 6
	for i, item := range items {
		row := startRow + i
		itemMap, _ := item.(map[string]interface{})

		// 设置单元格值
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), itemMap["序号"])
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), itemMap["品牌"])
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), itemMap["区域"])
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), itemMap["系统说明"])
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), itemMap["单位"])
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), itemMap["工程量"])
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), itemMap["预算价"])
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), itemMap["单项预算合价"])

		// 设置样式
		f.SetCellStyle(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("D%d", row), dataStyle)
		f.SetCellStyle(sheetName, fmt.Sprintf("E%d", row), fmt.Sprintf("F%d", row), dataStyle)
		f.SetCellStyle(sheetName, fmt.Sprintf("G%d", row), fmt.Sprintf("H%d", row), amountStyle)
	}

	// 添加总计行
	totalRow := startRow + len(items)
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", totalRow), "项目合计总价")
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", totalRow), "（不含增值税）")
	// 使用SetCellFormula方法设置公式，而不是excelize.Formula函数
	f.SetCellFormula(sheetName, fmt.Sprintf("H%d", totalRow), fmt.Sprintf("SUM(H%d:H%d)", startRow, totalRow-1))

	// 合并总计行
	// 修复MergeCell调用，使用正确的3参数格式
	f.MergeCell(sheetName, fmt.Sprintf("A%d", totalRow), fmt.Sprintf("C%d", totalRow))
	f.MergeCell(sheetName, fmt.Sprintf("D%d", totalRow), fmt.Sprintf("G%d", totalRow))

	// 设置总计行样式
	f.SetCellStyle(sheetName, fmt.Sprintf("A%d", totalRow), fmt.Sprintf("G%d", totalRow), headerStyle)
	f.SetCellStyle(sheetName, fmt.Sprintf("H%d", totalRow), fmt.Sprintf("H%d", totalRow), amountStyle)

	// 添加大写金额行
	capitalRow := totalRow + 1
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", capitalRow), "总价大写（元）")
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", capitalRow), "贰万贰仟肆佰伍拾玖元贰角整")
	f.MergeCell(sheetName, fmt.Sprintf("A%d", capitalRow), fmt.Sprintf("C%d", capitalRow))
	f.MergeCell(sheetName, fmt.Sprintf("D%d", capitalRow), fmt.Sprintf("H%d", capitalRow))
	f.SetCellStyle(sheetName, fmt.Sprintf("A%d", capitalRow), fmt.Sprintf("H%d", capitalRow), dataStyle)

	// 添加温馨提醒
	reminderRow := capitalRow + 2
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", reminderRow), "温馨提醒：")
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", reminderRow), "1、该报价为根据报价需求提供的方案报价，实际成交价以签约合同为准。 2、报价单仅供预算参考，具体内容以实际签订的合同为准。")
	f.MergeCell(sheetName, fmt.Sprintf("B%d", reminderRow), fmt.Sprintf("H%d", reminderRow))
	f.SetCellStyle(sheetName, fmt.Sprintf("A%d", reminderRow), fmt.Sprintf("H%d", reminderRow), dataStyle)

	// 动态调整数据行高，适应内容
	// for i := 6; i <= reminderRow; i++ {
	// 	// 对于每一行，根据内容长度动态调整行高
	// 	// 计算所需的行高（基础高度 + 内容长度/每行字符数 * 行高）
	// 	baseHeight := 20.0
	// 	// 获取该行的主要内容单元格（系统说明列D）
	// 	contentCell := fmt.Sprintf("D%d", i)
	// 	content, _ := f.GetCellValue(sheetName, contentCell)
	// 	// 计算所需行数（假设每行能显示30个字符）
	// 	lines := len(content) / 30
	// 	if len(content)%30 > 0 {
	// 		lines++
	// 	}
	// 	// 设置行高（最小20，最大120）
	// 	height := baseHeight * float64(lines)
	// 	if height < 20 {
	// 		height = 20
	// 	} else if height > 120 {
	// 		height = 120
	// 	}
	// 	f.SetRowHeight(sheetName, i, height)
	// }

	return nil
}

// processImages 处理图片
func (s *ExcelService) processImages(f *excelize.File, sheetName string, images []interface{}) error {
	for _, imgItem := range images {
		imgMap, ok := imgItem.(map[string]interface{})
		if !ok {
			continue
		}

		path, _ := imgMap["path"].(string)
		if path == "" {
			continue
		}

		// 暂时跳过图片处理，后续根据Excelize v2 API更新
		continue
	}

	return nil
}
