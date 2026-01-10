package main

import (
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/gin-gonic/gin"
	"office-export-server/internal/api"
	"office-export-server/internal/config"
)

func main() {
	// 解析命令行参数
	configFile := flag.String("config", "config.yaml", "path to config file")
	flag.Parse()

	// 加载配置
	if err := config.LoadConfig(*configFile); err != nil {
		log.Printf("Warning: failed to load config file, using default settings: %v", err)
	}

	// 设置Gin模式
	if config.GlobalConfig.Log.Level == "debug" {
		gin.SetMode(gin.DebugMode)
	} else {
		gin.SetMode(gin.ReleaseMode)
	}

	// 创建Gin路由引擎
	router := gin.Default()

	// 配置路由
	api.SetupRoutes(router)

	// 启动服务器
	serverAddr := fmt.Sprintf("%s:%d", config.GlobalConfig.Server.Host, config.GlobalConfig.Server.Port)
	log.Printf("Office Export Server is starting on %s", serverAddr)

	if err := router.Run(serverAddr); err != nil {
		log.Fatalf("Failed to start server: %v", err)
		os.Exit(1)
	}
}
