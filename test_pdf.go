package main

import (
	"fmt"
	"log"
	"os"

	"office-export-server/internal/model"
	"office-export-server/internal/service/export"
)

func main() {
	// 创建PDF导出服务
	pdfService := export.NewPDFService()

	// 构建测试请求
	req := &model.ExportRequest{
		Data: map[string]interface{}{
			"title": "全宅智能定制方案",
			"project": map[string]interface{}{
				"name":            "全宅智能定制方案20260112",
				"clientName":      "测试客户",
				"clientPhone":     "13800138000",
				"houseType":       "三室一厅",
				"address":         "北京市朝阳区",
				"serviceProvider": "xxx",
				"contact":         "xxx",
				"contactPhone":    "15625630782",
			},
			"products": []interface{}{
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
			},
		},
	}

	// 执行PDF导出
	data, err := pdfService.ExportPDF(req)
	if err != nil {
		log.Fatalf("PDF导出失败: %v", err)
	}

	// 保存PDF文件
	err = os.WriteFile("test_pdf_optimized.pdf", data, 0644)
	if err != nil {
		log.Fatalf("保存PDF文件失败: %v", err)
	}

	fmt.Println("PDF导出成功！文件已保存为 test_pdf_optimized.pdf")
}
