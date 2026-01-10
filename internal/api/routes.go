package api

import (
	"office-export-server/internal/api/handlers"
	"office-export-server/internal/service/export"
	"office-export-server/internal/service/template"

	"github.com/gin-gonic/gin"
)

// SetupRoutes 配置路由
func SetupRoutes(router *gin.Engine) {
	// 添加完整的CORS中间件
	router.Use(func(c *gin.Context) {
		c.Writer.Header().Set("Access-Control-Allow-Origin", "*")
		c.Writer.Header().Set("Access-Control-Allow-Credentials", "true")
		c.Writer.Header().Set("Access-Control-Allow-Headers", "Content-Type, Content-Length, Accept-Encoding, X-CSRF-Token, Authorization, accept, origin, Cache-Control, X-Requested-With")
		c.Writer.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS, GET, PUT, DELETE")

		if c.Request.Method == "OPTIONS" {
			c.AbortWithStatus(204)
			return
		}

		c.Next()
	})

	// 创建服务实例
	templateService := template.NewTemplateService()
	exportService := export.NewExportService(templateService)

	// 创建处理器实例
	exportHandler := handlers.NewExportHandler(exportService)
	templateHandler := handlers.NewTemplateHandler(templateService)

	// API分组
	api := router.Group("/api/v1")
	{
		// 导出相关路由
		export := api.Group("/export")
		{
			export.POST("/:type", exportHandler.ExportFile)
		}

		// 模板相关路由
		template := api.Group("/templates")
		{
			template.GET("", templateHandler.GetAllTemplates)
		}

		// 健康检查
		api.GET("/health", func(c *gin.Context) {
			c.JSON(200, gin.H{
				"status":  "ok",
				"message": "Office Export Server is running",
			})
		})
	}
}
