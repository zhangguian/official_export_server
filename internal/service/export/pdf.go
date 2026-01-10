package export

import (
	"bytes"
	"fmt"

	"office-export-server/internal/model"

	"github.com/jung-kurt/gofpdf"
)

// PDFService PDF导出服务
type PDFService struct{}

// NewPDFService 创建PDF导出服务实例
func NewPDFService() *PDFService {
	return &PDFService{}
}

// ExportPDF 导出PDF文件
func (s *PDFService) ExportPDF(req *model.ExportRequest) ([]byte, error) {
	// 创建新的PDF文档
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.AddPage()

	// 设置字体
	pdf.SetFont("Arial", "", 12)

	// 设置文档标题
	title, ok := req.Data["title"].(string)
	if ok && title != "" {
		pdf.SetFont("Arial", "B", 24)
		pdf.CellFormat(0, 10, title, "", 1, "C", false, 0, "")
		pdf.Ln(10)
		pdf.SetFont("Arial", "", 12)
	}

	// 处理正文内容
	content, ok := req.Data["content"].([]interface{})
	if ok {
		for _, item := range content {
			contentMap, ok := item.(map[string]interface{})
			if !ok {
				continue
			}

			// 根据内容类型处理
			contentType, ok := contentMap["type"].(string)
			if !ok {
				continue
			}

			switch contentType {
			case "paragraph":
				s.addParagraph(pdf, contentMap)
			case "table":
				s.addTable(pdf, contentMap)
			case "image":
				s.addImage(pdf, contentMap)
			}
		}
	}

	// 保存文档到缓冲区
	var buf bytes.Buffer
	err := pdf.Output(&buf)
	if err != nil {
		return nil, fmt.Errorf("failed to save pdf document: %v", err)
	}

	return buf.Bytes(), nil
}

// addParagraph 添加段落
func (s *PDFService) addParagraph(pdf *gofpdf.Fpdf, contentMap map[string]interface{}) {
	text, ok := contentMap["text"].(string)
	if !ok || text == "" {
		return
	}

	// 设置字体样式
	fontStyle := ""
	if bold, ok := contentMap["bold"].(bool); ok && bold {
		fontStyle += "B"
	}
	if italic, ok := contentMap["italic"].(bool); ok && italic {
		fontStyle += "I"
	}

	fontSize, ok := contentMap["font_size"].(float64)
	if !ok || fontSize <= 0 {
		fontSize = 12
	}

	pdf.SetFont("Arial", fontStyle, fontSize)

	// 添加段落
	pdf.MultiCell(0, 5, text, "", "L", false)
	pdf.Ln(5)

	// 恢复默认字体
	pdf.SetFont("Arial", "", 12)
}

// addTable 添加表格
func (s *PDFService) addTable(pdf *gofpdf.Fpdf, contentMap map[string]interface{}) {
	tableData, ok := contentMap["data"].(map[string]interface{})
	if !ok {
		return
	}

	// 获取表头和数据
	headers, ok := tableData["headers"].([]interface{})
	if !ok {
		return
	}

	rows, ok := tableData["rows"].([]interface{})
	if !ok {
		return
	}

	// 设置表格列宽
	colWidths, ok := tableData["col_widths"].([]interface{})
	if !ok || len(colWidths) == 0 {
		// 默认列宽
		colCount := 0
		for _, headerRow := range headers {
			headerCells, ok := headerRow.([]interface{})
			if ok && len(headerCells) > colCount {
				colCount = len(headerCells)
			}
		}
		colWidths = make([]interface{}, colCount)
		for i := range colWidths {
			colWidths[i] = 40.0
		}
	}

	// 转换列宽为float64切片
	widths := make([]float64, len(colWidths))
	for i, w := range colWidths {
		widths[i], _ = w.(float64)
	}

	// 绘制表头
	pdf.SetFont("Arial", "B", 12)
	pdf.SetFillColor(200, 220, 255)
	pdf.SetTextColor(0, 0, 0)
	pdf.SetLineWidth(0.3)

	for _, headerRow := range headers {
		headerCells, ok := headerRow.([]interface{})
		if !ok {
			continue
		}

		for i, cellItem := range headerCells {
			cellMap, ok := cellItem.(map[string]interface{})
			if !ok {
				continue
			}

			cellText, _ := cellMap["text"].(string)
			colSpan, _ := cellMap["col_span"].(float64)
			rowSpan, _ := cellMap["row_span"].(float64)

			if colSpan <= 0 {
				colSpan = 1
			}
			if rowSpan <= 0 {
				rowSpan = 1
			}

			// 计算合并后的宽度
			mergeWidth := 0.0
			for j := 0; j < int(colSpan); j++ {
				if i+j < len(widths) {
					mergeWidth += widths[i+j]
				}
			}

			pdf.CellFormat(mergeWidth, 7, cellText, "1", 0, "C", true, 0, "")
		}
		pdf.Ln(5)
	}

	// 绘制数据行
	pdf.SetFont("Arial", "", 11)
	pdf.SetFillColor(255, 255, 255)

	for _, rowItem := range rows {
		rowCells, ok := rowItem.([]interface{})
		if !ok {
			continue
		}

		for i, cellItem := range rowCells {
			cellMap, ok := cellItem.(map[string]interface{})
			if !ok {
				continue
			}

			cellText, _ := cellMap["text"].(string)
			colSpan, _ := cellMap["col_span"].(float64)
			rowSpan, _ := cellMap["row_span"].(float64)

			if colSpan <= 0 {
				colSpan = 1
			}
			if rowSpan <= 0 {
				rowSpan = 1
			}

			// 计算合并后的宽度
			mergeWidth := 0.0
			for j := 0; j < int(colSpan); j++ {
				if i+j < len(widths) {
					mergeWidth += widths[i+j]
				}
			}

			pdf.CellFormat(mergeWidth, 7, cellText, "1", 0, "L", false, 0, "")
		}
		pdf.Ln(5)
	}

	pdf.Ln(10)
}

// addImage 添加图片
func (s *PDFService) addImage(pdf *gofpdf.Fpdf, contentMap map[string]interface{}) {
	path, ok := contentMap["path"].(string)
	if !ok || path == "" {
		return
	}

	position, _ := contentMap["position"].(map[string]interface{})
	x, _ := position["x"].(float64)
	y, _ := position["y"].(float64)

	size, _ := contentMap["size"].(map[string]interface{})
	width, _ := size["width"].(float64)
	height, _ := size["height"].(float64)

	// 默认尺寸
	if width == 0 {
		width = 100
	}
	if height == 0 {
		height = 100
	}

	// 添加图片
	pdf.ImageOptions(path, x, y, width, height, false, gofpdf.ImageOptions{
		ReadDpi: true,
	}, 0, "")

	pdf.Ln(height + 10)
}
