package handlers

import (
	"net/http"

	"office-export-server/internal/model"
	"office-export-server/internal/service/export"
	"github.com/gin-gonic/gin"
)

// ExportHandler 导出处理器
type ExportHandler struct {
	exportService export.ExportService
}

// NewExportHandler 创建导出处理器实例
func NewExportHandler(exportService export.ExportService) *ExportHandler {
	return &ExportHandler{
		exportService: exportService,
	}
}

// ExportFile 处理文件导出请求
func (h *ExportHandler) ExportFile(c *gin.Context) {
	// 获取文件类型
	fileType := c.Param("type")
	if fileType == "" {
		c.JSON(http.StatusBadRequest, model.ErrorResponse{
			Code:    http.StatusBadRequest,
			Message: "file type is required",
		})
		return
	}

	// 绑定请求参数
	var req model.ExportRequest
	if err := c.ShouldBindJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, model.ErrorResponse{
			Code:    http.StatusBadRequest,
			Message: "invalid request parameters: " + err.Error(),
		})
		return
	}

	// 根据文件类型导出
	var fileBytes []byte
	var err error
	var contentType string
	var filename string

	switch fileType {
	case "excel":
		fileBytes, err = h.exportService.ExportExcel(&req)
		contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		filename = "export.xlsx"
	case "word":
		fileBytes, err = h.exportService.ExportWord(&req)
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
		filename = "export.docx"
	case "pdf":
		fileBytes, err = h.exportService.ExportPDF(&req)
		contentType = "application/pdf"
		filename = "export.pdf"
	default:
		c.JSON(http.StatusBadRequest, model.ErrorResponse{
			Code:    http.StatusBadRequest,
			Message: "unsupported file type",
		})
		return
	}

	if err != nil {
		c.JSON(http.StatusInternalServerError, model.ErrorResponse{
			Code:    http.StatusInternalServerError,
			Message: "failed to export file: " + err.Error(),
		})
		return
	}

	// 设置响应头
	c.Header("Content-Disposition", "attachment; filename=\""+filename+"\"")
	c.Header("Content-Type", contentType)
	c.Data(http.StatusOK, contentType, fileBytes)
}
