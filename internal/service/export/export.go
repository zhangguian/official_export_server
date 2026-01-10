package export

import (
	"office-export-server/internal/model"
	"office-export-server/internal/service/template"
)

// ExportService 导出服务接口
type ExportService interface {
	ExportExcel(req *model.ExportRequest) ([]byte, error)
	ExportWord(req *model.ExportRequest) ([]byte, error)
	ExportPDF(req *model.ExportRequest) ([]byte, error)
}

// exportService 导出服务实现
type exportService struct {
	excelService *ExcelService
	wordService  *WordService
	pdfService   *PDFService
}

// NewExportService 创建导出服务实例
func NewExportService(templateService template.TemplateService) ExportService {
	return &exportService{
		excelService: NewExcelService(templateService),
		wordService:  NewWordService(),
		pdfService:   NewPDFService(),
	}
}

// ExportExcel 导出Excel文件
func (s *exportService) ExportExcel(req *model.ExportRequest) ([]byte, error) {
	return s.excelService.ExportExcel(req)
}

// ExportWord 导出Word文件
func (s *exportService) ExportWord(req *model.ExportRequest) ([]byte, error) {
	return s.wordService.ExportWord(req)
}

// ExportPDF 导出PDF文件
func (s *exportService) ExportPDF(req *model.ExportRequest) ([]byte, error) {
	return s.pdfService.ExportPDF(req)
}
