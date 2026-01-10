package template

import (
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"

	"office-export-server/internal/config"
	"office-export-server/internal/model"
)

// TemplateService 模板服务接口
type TemplateService interface {
	LoadTemplate(templateID string, fileType string) ([]byte, error)
	GetAllTemplates() ([]model.TemplateInfo, error)
	GetTemplatePath(templateID string, fileType string) (string, error)
}

// templateService 模板服务实现
type templateService struct {
	templateDir string
}

// NewTemplateService 创建模板服务实例
func NewTemplateService() TemplateService {
	return &templateService{
		templateDir: config.GlobalConfig.Template.Path,
	}
}

// LoadTemplate 加载模板文件
func (s *templateService) LoadTemplate(templateID string, fileType string) ([]byte, error) {
	filePath, err := s.GetTemplatePath(templateID, fileType)
	if err != nil {
		return nil, err
	}

	data, err := ioutil.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to read template file: %v", err)
	}

	return data, nil
}

// GetAllTemplates 获取所有模板信息
func (s *templateService) GetAllTemplates() ([]model.TemplateInfo, error) {
	var templates []model.TemplateInfo

	// 遍历所有模板类型目录
	for _, fileType := range []string{"excel", "word", "pdf"} {
		typeDir := filepath.Join(s.templateDir, fileType)
		if _, err := os.Stat(typeDir); os.IsNotExist(err) {
			continue
		}

		// 读取目录下所有文件
		files, err := ioutil.ReadDir(typeDir)
		if err != nil {
			return nil, fmt.Errorf("failed to read template directory: %v", err)
		}

		// 处理每个模板文件
		for _, file := range files {
			if file.IsDir() {
				continue
			}

			fileName := file.Name()
			// 过滤掉Excel临时文件（以~$开头）
			if strings.HasPrefix(fileName, "~") {
				continue
			}
			templateID := strings.TrimSuffix(fileName, filepath.Ext(fileName))

			templateInfo := model.TemplateInfo{
				ID:          templateID,
				Name:        fileName,
				Description: fmt.Sprintf("%s template", fileType),
				Type:        fileType,
				Path:        filepath.Join(typeDir, fileName),
			}

			templates = append(templates, templateInfo)
		}
	}

	return templates, nil
}

// GetTemplatePath 获取模板文件路径
func (s *templateService) GetTemplatePath(templateID string, fileType string) (string, error) {
	// 根据文件类型确定文件扩展名
	extension := ""
	switch fileType {
	case "excel":
		extension = ".xlsx"
	case "word":
		extension = ".docx"
	case "pdf":
		extension = ".pdf"
	default:
		return "", fmt.Errorf("unsupported file type: %s", fileType)
	}

	// 构建模板文件路径
	filePath := filepath.Join(s.templateDir, fileType, templateID+extension)

	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return "", fmt.Errorf("template not found: %s", filePath)
	}

	return filePath, nil
}
