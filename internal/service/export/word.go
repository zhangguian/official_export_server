package export

import (
	"fmt"

	"office-export-server/internal/model"
)

// WordService Word导出服务
type WordService struct{}

// NewWordService 创建Word导出服务实例
func NewWordService() *WordService {
	return &WordService{}
}

// ExportWord 导出Word文件（暂时简化实现）
func (s *WordService) ExportWord(req *model.ExportRequest) ([]byte, error) {
	return nil, fmt.Errorf("Word export feature is currently under development, please try Excel or PDF export instead")
}
