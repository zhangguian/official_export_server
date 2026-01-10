# 多Sheet页功能使用说明

## 功能介绍

Office Export Server 支持导出包含多个Sheet页的Excel文件。通过在请求数据中添加`sheets`数组，可以指定多个Sheet页的名称和数据，服务会自动创建对应的Sheet页并填充数据。

## 请求格式

### 基本URL
```
http://localhost:8080/api/v1/export/excel
```

### 请求方法
```
POST
```

### 请求头
```
Content-Type: application/json
```

## 数据结构

### 基本结构

```json
{
  "template_id": "default",
  "data_type": "excel",
  "data": {
    "sheets": [
      {
        "name": "Sheet1",
        "items": [
          // 第一Sheet的数据
        ]
      },
      {
        "name": "Sheet2",
        "items": [
          // 第二Sheet的数据
        ]
      }
    ]
  }
}
```

### 字段说明

| 字段名 | 类型 | 必填 | 描述 |
|-------|------|------|------|
| template_id | string | 否 | 模板ID，默认为"default" |
| data_type | string | 是 | 必须为"excel" |
| data | object | 是 | 包含sheets数组的对象 |
| data.sheets | array | 是 | 包含多个Sheet页配置的数组 |
| data.sheets[].name | string | 是 | Sheet页名称 |
| data.sheets[].items | array | 否 | 当前Sheet页的数据数组 |

## 示例代码

### 示例1：使用PowerShell调用API

```powershell
# 创建请求数据
$requestData = @{
    template_id = "default"
    data_type = "excel"
    data = @{
        sheets = @(
            @{
                name = "预算汇总表"
                items = @(
                    @{
                        "序号" = 1
                        "品牌" = "小米"
                        "区域" = "全屋智能主控系统"
                        "系统说明" = "1、AI智能语音、自定义设备各种场景..."
                        "单位" = "项"
                        "工程量" = 1
                        "预算价" = 1928
                        "单项预算合价" = 1928
                    }
                )
            },
            @{
                name = "报价单"
                items = @(
                    @{
                        "品名" = "大班台"
                        "规格" = "2400*2000*750"
                        "材质说明" = "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面"
                        "颜色" = "黑色"
                        "数量" = 2
                        "单价" = 10141.00
                        "总价" = 20282.00
                        "备注" = ""
                    }
                )
            }
        )
    }
}

# 转换为JSON
$jsonData = $requestData | ConvertTo-Json -Depth 10

# 发送请求
Invoke-WebRequest -Uri "http://localhost:8080/api/v1/export/excel" -Method POST -ContentType "application/json" -Body $jsonData -OutFile "test_multi_sheet.xlsx"
```

### 示例2：使用Go语言调用API

```go
package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
)

func main() {
	url := "http://localhost:8080/api/v1/export/excel"

	// 准备多sheet数据
	data := map[string]interface{}{
		"template_id": "default",
		"data_type":   "excel",
		"data": map[string]interface{}{
			"sheets": []map[string]interface{}{
				{
					"name": "预算汇总表",
					"items": []map[string]interface{}{
						{
							"序号":          1,
							"品牌":          "小米",
							"区域":          "全屋智能主控系统",
							"系统说明":      "1、AI智能语音、自定义设备各种场景...",
							"单位":          "项",
							"工程量":        1,
							"预算价":        1928,
							"单项预算合价":  1928,
						},
					},
				},
				{
					"name": "报价单",
					"items": []map[string]interface{}{
						{
							"品名":      "大班台",
							"规格":      "2400*2000*750",
							"材质说明":  "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
							"颜色":      "黑色",
							"数量":      2,
							"单价":      10141.00,
							"总价":      20282.00,
							"备注":      "",
						},
					},
				},
			},
		},
	}

	// 编码为JSON
	jsonData, err := json.Marshal(data)
	if err != nil {
		fmt.Printf("JSON编码失败: %v\n", err)
		return
	}

	// 发送POST请求
	resp, err := http.Post(url, "application/json", bytes.NewBuffer(jsonData))
	if err != nil {
		fmt.Printf("HTTP请求失败: %v\n", err)
		return
	}
	defer resp.Body.Close()

	// 检查响应状态
	if resp.StatusCode == http.StatusOK {
		// 读取响应体
		body, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			fmt.Printf("读取响应失败: %v\n", err)
			return
		}

		// 保存为文件
		err = ioutil.WriteFile("test_multi_sheet.xlsx", body, 0644)
		if err != nil {
			fmt.Printf("保存文件失败: %v\n", err)
			return
		}

		fmt.Println("多sheet页导出成功！")
	} else {
		// 读取错误响应
		errBody, _ := ioutil.ReadAll(resp.Body)
		fmt.Printf("导出失败，状态码: %d, 错误信息: %s\n", resp.StatusCode, string(errBody))
	}
}
```

### 示例3：使用JavaScript调用API

```javascript
async function exportExcel() {
  const url = 'http://localhost:8080/api/v1/export/excel';
  const data = {
    template_id: 'default',
    data_type: 'excel',
    data: {
      sheets: [
        {
          name: '预算汇总表',
          items: [
            {
              '序号': 1,
              '品牌': '小米',
              '区域': '全屋智能主控系统',
              '系统说明': '1、AI智能语音、自定义设备各种场景...',
              '单位': '项',
              '工程量': 1,
              '预算价': 1928,
              '单项预算合价': 1928
            }
          ]
        },
        {
          name: '报价单',
          items: [
            {
              '品名': '大班台',
              '规格': '2400*2000*750',
              '材质说明': '环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面',
              '颜色': '黑色',
              '数量': 2,
              '单价': 10141.00,
              '总价': 20282.00,
              '备注': ''
            }
          ]
        }
      ]
    }
  };

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(data)
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'test_multi_sheet.xlsx';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      console.log('导出成功！');
    } else {
      const error = await response.json();
      console.error('导出失败:', error);
    }
  } catch (error) {
    console.error('导出失败:', error);
  }
}

// 调用函数
exportExcel();
```

## 注意事项

1. **Sheet名称唯一性**：每个Sheet页的名称应该唯一，避免重复。

2. **模板兼容性**：多Sheet页功能适用于所有模板，服务会为每个Sheet页应用相同的模板样式。

3. **数据结构一致性**：同一模板的不同Sheet页，数据结构应该保持一致，确保数据能正确填充到模板中。

4. **性能考虑**：创建大量Sheet页可能会影响导出性能，建议根据实际需求合理规划Sheet数量。

5. **Sheet数量限制**：Excel文件本身对Sheet数量有一定限制（通常为1048576个），但实际使用中建议控制在合理范围内。

6. **图片处理**：如果需要在多个Sheet页中插入图片，每个Sheet页的数据中需要包含对应的图片信息。

## 常见问题

### Q: 导出的Excel文件只有一个Sheet页？
A: 请检查请求数据中是否正确包含`sheets`数组，以及数组中的每个Sheet对象是否包含`name`字段。

### Q: Sheet名称不正确？
A: 请确保每个Sheet对象中的`name`字段是字符串类型，且不为空。

### Q: 某个Sheet页的数据没有显示？
A: 请检查该Sheet对象中的`items`数组是否正确，以及数据结构是否与模板预期一致。

### Q: 导出失败，返回400错误？
A: 请检查请求数据格式是否正确，特别是JSON格式是否规范，是否缺少必填字段。

## 版本要求

- Office Export Server v1.0.0 及以上版本支持多Sheet页功能。

## 测试工具

服务提供了一个测试脚本，可以用于验证多Sheet页功能：

```bash
# 运行测试脚本
go run test_multi_sheet.go
```

脚本会生成一个包含两个Sheet页的Excel文件 `test_multi_sheet.xlsx`，可以用Excel打开查看效果。

## 最佳实践

1. **合理规划Sheet数量**：根据实际需求规划Sheet数量，避免创建过多不必要的Sheet页。

2. **统一Sheet命名规范**：使用清晰、有意义的Sheet名称，便于用户理解和使用。

3. **保持数据结构一致性**：同一模板的不同Sheet页，数据结构应该保持一致，便于模板设计和维护。

4. **测试简单数据**：在正式使用前，先用简单数据测试多Sheet页功能，确保格式正确。

5. **监控性能**：对于大量数据的多Sheet页导出，建议监控导出性能，必要时进行优化。

## 结论

多Sheet页功能为Excel导出提供了更大的灵活性，支持用户根据需要创建多个Sheet页，每个Sheet页包含独立的数据。通过遵循本文档的规范和建议，可以轻松使用多Sheet页功能，导出符合需求的Excel文件。