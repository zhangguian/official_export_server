package handlers

import (
	"net/http"

	"github.com/gin-gonic/gin"
	"office-export-server/internal/model"
	"office-export-server/internal/service/template"
)

// TemplateHandler 模板处理器
type TemplateHandler struct {
	templateService template.TemplateService
}

// NewTemplateHandler 创建模板处理器实例
func NewTemplateHandler(templateService template.TemplateService) *TemplateHandler {
	return &TemplateHandler{
		templateService: templateService,
	}
}

// GetAllTemplates 获取所有模板信息
func (h *TemplateHandler) GetAllTemplates(c *gin.Context) {
	templates, err := h.templateService.GetAllTemplates()
	if err != nil {
		c.JSON(http.StatusInternalServerError, model.ErrorResponse{
			Code:    http.StatusInternalServerError,
			Message: "failed to get templates: " + err.Error(),
		})
		return
	}

	c.JSON(http.StatusOK, model.SuccessResponse{
		Code:    http.StatusOK,
		Message: "success",
		Data:    templates,
	})
}
