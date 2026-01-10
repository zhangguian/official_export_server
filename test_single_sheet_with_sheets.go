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

	// 准备单sheet数据，使用sheets数组格式
	data := map[string]interface{}{
		"template_id": "default",
		"data_type":   "excel",
		"data": map[string]interface{}{
			"sheets": []map[string]interface{}{
				{
					"name": "预算汇总表",
					"items": []map[string]interface{}{
						{
							"序号":     1,
							"品牌":     "小米",
							"区域":     "全屋智能主控系统",
							"系统说明":   "1、AI智能语音、自定义设备各种场景...",
							"单位":     "项",
							"工程量":    1,
							"预算价":    1928,
							"单项预算合价": 1928,
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
		err = ioutil.WriteFile("test_single_sheet_with_sheets.xlsx", body, 0644)
		if err != nil {
			fmt.Printf("保存文件失败: %v\n", err)
			return
		}

		fmt.Println("使用sheets数组的单sheet页导出成功！文件已保存为 test_single_sheet_with_sheets.xlsx")
	} else {
		// 读取错误响应
		errBody, _ := ioutil.ReadAll(resp.Body)
		fmt.Printf("导出失败，状态码: %d, 错误信息: %s\n", resp.StatusCode, string(errBody))
	}
}
