# Office Export Server API使用文档

## 项目概述

Office Export Server是一个基于Go语言开发的Office文件导出服务，支持导出Word、Excel和PDF文件。该服务提供了RESTful API接口，允许前端或其他服务发送请求，生成包含指定数据和格式的Office文件。

## API基本信息

### 基础URL
```
http://localhost:8080/api/v1
```

### 认证方式
当前版本暂不支持认证，所有API均可匿名访问。

### 支持的文件类型
- Excel (.xlsx)
- Word (.docx) - 开发中
- PDF (.pdf) - 开发中

## 导出Excel API

### 请求方法和端点
```
POST /export/excel
```

### 请求参数

| 参数名 | 类型 | 必填 | 位置 | 描述 |
|-------|------|------|------|------|
| template_id | string | 否 | body | 模板ID，默认为"default" |
| data_type | string | 是 | body | 数据类型，必须为"excel" |
| data | object | 是 | body | 导出数据，包含items数组和其他可选字段 |

### 数据结构

```json
{
  "template_id": "default",
  "data_type": "excel",
  "data": {
    "sheet_name": "预算汇总表",
    "items": [
      {
        "序号": 1,
        "品牌": "小米",
        "区域": "全屋智能主控系统",
        "系统说明": "1、AI智能语音、自定义设备各种场景...",
        "单位": "项",
        "工程量": 1,
        "预算价": 1928,
        "单项预算合价": 1928
      }
    ]
  }
}
```

### 响应格式

#### 成功响应
- 状态码：200 OK
- 响应头：`Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- 响应体：Excel文件二进制数据

#### 错误响应
- 状态码：400 Bad Request 或 500 Internal Server Error
- 响应体：
  ```json
  {
    "code": 400,
    "message": "invalid request parameters: Key: 'ExportRequest.DataType' Error:Field validation for 'DataType' failed on the 'required' tag"
  }
  ```

### 示例请求

#### 使用PowerShell
```powershell
Invoke-WebRequest -Uri http://localhost:8080/api/v1/export/excel -Method POST -ContentType 'application/json' -Body '{"template_id":"default","data_type":"excel","data":{"items":[{"序号":1,"品牌":"小米","区域":"全屋智能主控系统","系统说明":"1、AI智能语音、自定义设备各种场景（回家、离家、会客、就餐、休闲、阅读等模式），完美实现智能化体验。 2、智能品类包括：智能灯光、智能遮阳、智能空调，智能安防等； 3、最大优势及亮点 \"无缝接入米家APP、AI智能语音控制、轻成本、轻设计、轻方案、轻对接、轻落地、轻维护\"； 4、可以根据所需的智能开关与空调语音小助手进行DIY定制。","单位":"项","工程量":1,"预算价":1928,"单项预算合价":1928}]}}' -OutFile test_export.xlsx
```

#### 使用cURL
```bash
curl -X POST -H "Content-Type: application/json" -d '{"template_id":"default","data_type":"excel","data":{"items":[{"序号":1,"品牌":"小米","区域":"全屋智能主控系统","系统说明":"1、AI智能语音、自定义设备各种场景（回家、离家、会客、就餐、休闲、阅读等模式），完美实现智能化体验。 2、智能品类包括：智能灯光、智能遮阳、智能空调，智能安防等； 3、最大优势及亮点 \"无缝接入米家APP、AI智能语音控制、轻成本、轻设计、轻方案、轻对接、轻落地、轻维护\"； 4、可以根据所需的智能开关与空调语音小助手进行DIY定制。","单位":"项","工程量":1,"预算价":1928,"单项预算合价":1928}]}}' http://localhost:8080/api/v1/export/excel -o test_export.xlsx
```

## 导出Word API

### 请求方法和端点
```
POST /export/word
```

### 说明
该功能目前处于开发中，暂不支持使用。

## 导出PDF API

### 请求方法和端点
```
POST /export/pdf
```

### 说明
该功能目前处于开发中，暂不支持使用。

## 模板管理API

### 获取模板列表
```
GET /templates/:type
```

#### 参数
- `type`: 文件类型，可选值：`excel`、`word`、`pdf`

#### 响应示例
```json
[
  {
    "id": "default",
    "name": "默认模板",
    "type": "excel",
    "path": "templates/excel/default.xlsx"
  },
  {
    "id": "quote",
    "name": "报价单模板",
    "type": "excel",
    "path": "templates/excel/quote.xlsx"
  }
]
```

## 错误码说明

| 错误码 | 描述 |
|-------|------|
| 400 | 请求参数错误，如缺少必填字段或字段类型错误 |
| 404 | 模板不存在 |
| 500 | 服务器内部错误，如模板文件损坏或处理失败 |

## 最佳实践

1. **模板管理**：
   - 建议在 `templates/excel` 目录下放置不同的模板文件，用于不同场景的导出需求
   - 模板文件名即为模板ID，例如 `default.xlsx` 对应的模板ID为 `default`

2. **数据格式**：
   - 确保发送的数据格式与模板预期的格式一致
   - 对于复杂的模板，建议先测试简单数据，确保格式正确后再发送完整数据

3. **错误处理**：
   - 客户端应处理不同的HTTP状态码，并根据错误信息提示用户
   - 对于500错误，建议记录详细日志，方便排查问题

4. **性能考虑**：
   - 对于大量数据的导出，建议分批处理或优化数据结构
   - 图片导出可能会增加响应时间，建议优化图片大小和数量

## 示例代码

### 前端React示例

```typescript
import React from 'react';

const ExportButton = () => {
  const handleExport = async () => {
    const data = {
      template_id: "default",
      data_type: "excel",
      data: {
        sheet_name: "预算汇总表",
        items: [
          {
            "序号": 1,
            "品牌": "小米",
            "区域": "全屋智能主控系统",
            "系统说明": "1、AI智能语音、自定义设备各种场景...",
            "单位": "项",
            "工程量": 1,
            "预算价": 1928,
            "单项预算合价": 1928
          }
        ]
      }
    };

    try {
      const response = await fetch('http://localhost:8080/api/v1/export/excel', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(data),
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'export.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      } else {
        const error = await response.json();
        alert(`导出失败：${error.message}`);
      }
    } catch (error) {
      console.error('导出失败:', error);
      alert('导出失败，请稍后重试');
    }
  };

  return (
    <button onClick={handleExport}>
      导出Excel
    </button>
  );
};

export default ExportButton;
```

### 后端Go示例

```go
package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"net/http"
)

func main() {
	url := "http://localhost:8080/api/v1/export/excel"

	data := map[string]interface{}{
		"template_id": "default",
		"data_type":   "excel",
		"data": map[string]interface{}{
			"sheet_name": "预算汇总表",
			"items": []map[string]interface{}{
				{
					"序号":            1,
					"品牌":            "小米",
					"区域":            "全屋智能主控系统",
					"系统说明":        "1、AI智能语音、自定义设备各种场景...",
					"单位":            "项",
					"工程量":          1,
					"预算价":          1928,
					"单项预算合价":    1928,
				},
			},
		},
	}

	jsonData, err := json.Marshal(data)
	if err != nil {
		fmt.Println("JSON编码失败:", err)
		return
	}

	resp, err := http.Post(url, "application/json", bytes.NewBuffer(jsonData))
	if err != nil {
		fmt.Println("HTTP请求失败:", err)
		return
	}
	defer resp.Body.Close()

	if resp.StatusCode == http.StatusOK {
		fmt.Println("导出成功")
		// 处理响应体，保存文件
	} else {
		fmt.Printf("导出失败，状态码: %d\n", resp.StatusCode)
		// 处理错误响应
	}
}
```

## 常见问题

1. **Q: 导出的Excel文件没有显示图片？**
   A: 请确保图片URL是可访问的，并且服务器有网络访问权限。图片会被下载到临时文件，然后插入到Excel中。

2. **Q: 数据行的高度没有自适应内容？**
   A: 当前版本使用固定行高，后续会优化为自动行高。您可以在模板中预设置合适的行高。

3. **Q: 如何添加新的模板？**
   A: 只需将模板文件放入对应的模板目录即可，例如 `templates/excel/新模板.xlsx`，对应的模板ID为 `新模板`。

4. **Q: 支持多Sheet页导出吗？**
   A: 当前版本支持单Sheet页导出，多Sheet页功能正在开发中。

## 版本历史

### v1.0.0 (2026-01-10)
- 初始版本，支持Excel导出
- 支持模板管理
- 支持图片导出
- 支持自定义数据格式

## 联系方式

如有问题或建议，请联系开发团队。