package model

// ExportRequest 导出请求模型
type ExportRequest struct {
	TemplateID string                 `json:"template_id" binding:"required"`
	DataType   string                 `json:"data_type" binding:"required"`
	Data       map[string]interface{} `json:"data" binding:"required"`
}

// SheetData Excel Sheet数据模型
type SheetData struct {
	Name    string        `json:"name"`
	Headers [][]string    `json:"headers"`
	Rows    [][]interface{} `json:"rows"`
	Merges  []MergeRange  `json:"merges,omitempty"`
}

// MergeRange 合并单元格范围
type MergeRange struct {
	StartRow int `json:"start_row"`
	StartCol int `json:"start_col"`
	EndRow   int `json:"end_row"`
	EndCol   int `json:"end_col"`
}

// ImageData 图片数据模型
type ImageData struct {
	Path     string     `json:"path"`
	Position Position   `json:"position"`
	Size     ImageSize  `json:"size,omitempty"`
}

// Position 位置坐标
type Position struct {
	X float64 `json:"x"`
	Y float64 `json:"y"`
}

// ImageSize 图片尺寸
type ImageSize struct {
	Width  float64 `json:"width"`
	Height float64 `json:"height"`
}

// TableCell PDF表格单元格
type TableCell struct {
	Text     string `json:"text"`
	ColSpan  int    `json:"col_span,omitempty"`
	RowSpan  int    `json:"row_span,omitempty"`
}

// TableData PDF表格数据
type TableData struct {
	Headers [][]TableCell `json:"headers"`
	Rows    [][]TableCell `json:"rows"`
}
