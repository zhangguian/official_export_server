package config

import (
	"io/ioutil"
	"log"

	"gopkg.in/yaml.v3"
)

// Config 服务配置结构体
type Config struct {
	Server struct {
		Port int    `yaml:"port"`
		Host string `yaml:"host"`
	} `yaml:"server"`
	Template struct {
		Path string `yaml:"path"`
	} `yaml:"template"`
	Log struct {
		Level string `yaml:"level"`
	} `yaml:"log"`
}

// GlobalConfig 全局配置变量
var GlobalConfig Config

// LoadConfig 从YAML文件加载配置
func LoadConfig(filePath string) error {
	data, err := ioutil.ReadFile(filePath)
	if err != nil {
		log.Printf("Failed to read config file: %v", err)
		return err
	}

	if err := yaml.Unmarshal(data, &GlobalConfig); err != nil {
		log.Printf("Failed to parse config file: %v", err)
		return err
	}

	// 设置默认值
	if GlobalConfig.Server.Port == 0 {
		GlobalConfig.Server.Port = 8080
	}
	if GlobalConfig.Server.Host == "" {
		GlobalConfig.Server.Host = "0.0.0.0"
	}
	if GlobalConfig.Template.Path == "" {
		GlobalConfig.Template.Path = "./templates"
	}
	if GlobalConfig.Log.Level == "" {
		GlobalConfig.Log.Level = "info"
	}

	return nil
}
