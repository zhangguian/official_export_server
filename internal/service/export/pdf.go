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
	// 创建新的PDF文档，使用横向布局
	pdf := gofpdf.New("L", "mm", "A4", "")
	pdf.AddPage()

	// 添加中文字体支持
	// 注册中文字体（使用阿里巴巴普惠体）
	// 使用相对于当前工作目录的正确路径
	pdf.AddUTF8Font("AlibabaPuHuiTi", "", "templates/fonts/Alibaba_PuHuiTi_2.0_55_Regular_55_Regular.ttf")
	pdf.AddUTF8Font("AlibabaPuHuiTi", "B", "templates/fonts/Alibaba_PuHuiTi_2.0_75_SemiBold_75_SemiBold.ttf")
	pdf.AddUTF8Font("AlibabaPuHuiTi", "I", "templates/fonts/Alibaba_PuHuiTi_2.0_55_Regular_55_Regular.ttf")

	// 设置默认字体为中文字体
	pdf.SetFont("AlibabaPuHuiTi", "", 12)

	// 获取项目数据
	projectData, _ := req.Data["project"].(map[string]interface{})

	// 顶部Logo和标题区域
	pdf.SetXY(10, 10)
	pdf.SetFont("AlibabaPuHuiTi", "B", 14)
	pdf.CellFormat(100, 10, "ORVIBO欧瑞博", "", 0, "L", false, 0, "")
	pdf.SetXY(257, 10)
	pdf.SetFont("AlibabaPuHuiTi", "", 10)
	pdf.CellFormat(30, 10, "5G时代全宅智能 就选欧瑞博", "", 1, "R", false, 0, "")

	// 主视觉图片 - 调整尺寸，避免占用过多页面空间
	if imagePath, ok := projectData["coverImage"].(string); ok && imagePath != "" {
		pdf.ImageOptions(imagePath, 10, 20, 277, 70, false, gofpdf.ImageOptions{ReadDpi: true}, 0, "")
	} else {
		// 绘制一个占位矩形
		pdf.SetFillColor(200, 200, 200)
		pdf.Rect(10, 20, 277, 70, "F")
		pdf.SetFillColor(255, 255, 255)
		pdf.SetTextColor(0, 0, 0)
		pdf.SetFont("AlibabaPuHuiTi", "B", 18)
		pdf.CellFormat(277, 70, "全宅智能定制方案", "", 1, "CM", false, 0, "")
	}

	// 项目标题
	title, ok := req.Data["title"].(string)
	if !ok || title == "" {
		title = "全宅智能定制方案"
	}
	pdf.SetXY(10, 95)
	pdf.SetFont("AlibabaPuHuiTi", "B", 16)
	pdf.CellFormat(277, 10, title, "", 1, "C", false, 0, "")

	// 项目基本信息表格 - 紧跟在标题下方
	pdf.SetXY(10, 105)
	pdf.SetFont("AlibabaPuHuiTi", "B", 10)
	pdf.SetFillColor(240, 240, 240)

	// 项目信息字段
	infoFields := []struct {
		label string
		value string
	}{
		{"项目名称", "全宅智能定制方案20260112"},
		{"客户名称", ""},
		{"客户电话", ""},
		{"户型", "一室一厅"},
		{"地址", ""},
		{"服务商", "xxx"},
		{"联系人", "xxx"},
		{"联系电话", "15625630782"},
	}

	// 绘制项目信息表格
	pdf.SetLineWidth(0.2)
	for i, field := range infoFields {
		pdf.CellFormat(40, 7, field.label, "1", 0, "R", true, 0, "")
		pdf.CellFormat(55, 7, field.value, "1", 0, "L", false, 0, "")
		if (i+1)%3 == 0 {
			pdf.Ln(7)
		}
	}
	// 确保表格绘制完成后正确换行
	if len(infoFields)%3 != 0 {
		pdf.Ln(7)
	}

	// 产品清单标题 - 调整位置，紧跟在项目信息表格下方
	pdf.Ln(5) // 减少换行距离
	pdf.SetFont("AlibabaPuHuiTi", "B", 14)
	pdf.CellFormat(277, 10, "产品清单", "", 1, "L", false, 0, "")

	// 产品表格
	pdf.SetFont("AlibabaPuHuiTi", "B", 10)
	pdf.SetFillColor(200, 220, 255)
	pdf.SetTextColor(0, 0, 0)
	pdf.SetLineWidth(0.3)

	// 产品表头
	headers := []string{"产品", "单价", "数量", "金额", "产品说明"}
	// 调整列宽，充分利用横向页面宽度（A4横向宽度为297mm，左右边距各10mm，可用宽度277mm）
	colWidths := []float64{70, 35, 25, 35, 112}

	// 绘制表头
	for i, header := range headers {
		pdf.CellFormat(colWidths[i], 10, header, "1", 0, "C", true, 0, "")
	}
	pdf.Ln(10)

	// 产品数据
	products, ok := req.Data["products"].([]interface{})
	if !ok {
		// 使用示例数据
		products = []interface{}{
			map[string]interface{}{
				"name":        "MixSwitch智能开关(三开 雅灰)",
				"price":       "369.00",
				"quantity":    "1.00",
				"amount":      "¥369.00",
				"description": "MixSwitch 超级智能开关 采用MixPad家族式大按键设计，支持独创MixCtrl技术，按键功能自由定义；V0级防火PC材质，安全耐用。",
			},
			map[string]interface{}{
				"name":        "智能摄像机",
				"price":       "249.00",
				"quantity":    "1.00",
				"amount":      "¥249.00",
				"description": "智能摄像机 拥有355°高清广角、超强夜视及声音侦测告警功能，清晰的双向语音通话，让你不错过家人的任何重要时刻。",
			},
		}
	}

	// 绘制产品行
	pdf.SetFont("AlibabaPuHuiTi", "", 9)

	for i, productItem := range products {
		product, ok := productItem.(map[string]interface{})
		if !ok {
			continue
		}

		name, _ := product["name"].(string)
		price, _ := product["price"].(string)
		quantity, _ := product["quantity"].(string)
		amountStr, _ := product["amount"].(string)
		description, _ := product["description"].(string)

		// 交替行背景色
		if i%2 == 0 {
			pdf.SetFillColor(245, 250, 255)
		} else {
			pdf.SetFillColor(255, 255, 255)
		}

		// 计算行高
		lineHeight := 7.0
		descLines := 1
		if len(description) > 0 {
			descWidth := pdf.GetStringWidth(description)
			if descWidth > colWidths[4] {
				descLines = int(descWidth/colWidths[4]) + 1
			}
		}
		totalHeight := float64(descLines) * lineHeight

		// 简化的表格行绘制
		// 产品名称
		pdf.MultiCell(colWidths[0], lineHeight, name, "1", "L", true)
		// 单价
		pdf.SetXY(pdf.GetX()-colWidths[0]+colWidths[0], pdf.GetY()-lineHeight)
		pdf.MultiCell(colWidths[1], lineHeight, price, "1", "R", true)
		// 数量
		pdf.SetXY(pdf.GetX()-colWidths[1]+colWidths[0]+colWidths[1], pdf.GetY()-lineHeight)
		pdf.MultiCell(colWidths[2], lineHeight, quantity, "1", "C", true)
		// 金额
		pdf.SetXY(pdf.GetX()-colWidths[2]+colWidths[0]+colWidths[1]+colWidths[2], pdf.GetY()-lineHeight)
		pdf.MultiCell(colWidths[3], lineHeight, amountStr, "1", "R", true)
		// 产品说明（支持多行）
		pdf.SetXY(pdf.GetX()-colWidths[3]+colWidths[0]+colWidths[1]+colWidths[2]+colWidths[3], pdf.GetY()-lineHeight)
		pdf.MultiCell(colWidths[4], lineHeight, description, "1", "L", true)

		// 确保行高一致
		currentY := pdf.GetY()
		expectedY := pdf.GetY() - lineHeight + totalHeight
		if currentY < expectedY {
			pdf.SetXY(10, expectedY)
		} else {
			pdf.Ln(0)
		}
	}

	// 绘制总计行
	pdf.SetFont("AlibabaPuHuiTi", "B", 10)
	pdf.SetFillColor(200, 220, 255)

	// 总计行
	pdf.CellFormat(colWidths[0]+colWidths[1]+colWidths[2], 10, "总计", "1", 0, "R", true, 0, "")
	pdf.CellFormat(colWidths[3], 10, "¥6192.00", "1", 0, "R", true, 0, "")
	pdf.CellFormat(colWidths[4], 10, "", "1", 1, "L", true, 0, "")

	// 服务费行
	pdf.CellFormat(colWidths[0]+colWidths[1]+colWidths[2], 10, "服务费", "1", 0, "R", true, 0, "")
	pdf.CellFormat(colWidths[3], 10, "¥1238.40", "1", 0, "R", true, 0, "")
	pdf.CellFormat(colWidths[4], 10, "", "1", 1, "L", true, 0, "")

	// 最终总计行
	pdf.SetFont("AlibabaPuHuiTi", "B", 12)
	pdf.CellFormat(colWidths[0]+colWidths[1]+colWidths[2], 10, "总计", "1", 0, "R", true, 0, "")
	pdf.CellFormat(colWidths[3], 10, "¥7430.40", "1", 0, "R", true, 0, "")
	pdf.CellFormat(colWidths[4], 10, "", "1", 1, "L", true, 0, "")

	// 设置页脚回调函数，实现自动页码
	pdf.SetFooterFunc(func() {
		pdf.SetFont("AlibabaPuHuiTi", "", 8)
		pdf.SetXY(10, -10)
		pdf.CellFormat(200, 5, "全宅智能定制方案20260112", "", 0, "L", false, 0, "")
		pdf.SetXY(240, -10)
		pdf.CellFormat(20, 5, fmt.Sprintf("%d / %d", pdf.PageNo(), pdf.PageCount()), "", 0, "C", false, 0, "")
		pdf.SetXY(270, -10)
		pdf.CellFormat(20, 5, "xxx 15625630782", "", 1, "R", false, 0, "")
	})

	// 保存文档到缓冲区
	var buf bytes.Buffer
	err := pdf.Output(&buf)
	if err != nil {
		return nil, fmt.Errorf("failed to save pdf document: %v", err)
	}

	return buf.Bytes(), nil
}
