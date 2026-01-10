import { useState, useEffect } from 'react'
import axios from 'axios'
import './App.css'

interface TemplateInfo {
  id: string
  name: string
  description: string
  type: string
  path: string
}

interface ExportFormData {
  title: string
  templateId: string
  fileName: string
}

function App() {
  const [formData, setFormData] = useState<ExportFormData>({
    title: '测试报告',
    templateId: 'default',
    fileName: 'export'
  })

  const [templates, setTemplates] = useState<TemplateInfo[]>([])
  const [loading, setLoading] = useState(false)
  const [message, setMessage] = useState('')

  // 加载模板列表
  useEffect(() => {
    const fetchTemplates = async () => {
      try {
        const response = await axios.get('http://localhost:8080/api/v1/templates')
        setTemplates(response.data.data.filter((t: TemplateInfo) => t.type === 'excel'))
      } catch (error) {
        console.error('获取模板列表失败:', error)
        setMessage('获取模板列表失败，请检查服务是否正常运行')
      }
    }
    fetchTemplates()
  }, [])

  // 处理表单输入变化
  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target
    setFormData(prev => ({ ...prev, [name]: value }))
  }

  // 导出文件的通用函数
  const exportFile = async (fileType: 'excel' | 'word' | 'pdf') => {
    setLoading(true)
    setMessage('')

    try {
      let requestData: any = {
        template_id: formData.templateId,
        data_type: fileType,
        data: {
          title: formData.title
        }
      }

      // 如果是Excel文件，根据模板类型使用不同的数据结构
      if (fileType === 'excel') {
        // 报价单模板数据
        if (formData.templateId === 'quote') {
          requestData.data = {
            sheets: [
              {
                name: '报价单',
                title: '稻壳科技有限公司报价单',
                items: [
                  {
                    "品名": "大班台",
                    "规格": "2400*2000*750",
                    "材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
                    "颜色": "黑色",
                    "数量": 2,
                    "单价": 10141.00,
                    "总价": 20282.00,
                    "备注": ""
                  },
                  {
                    "品名": "文件柜",
                    "规格": "2400*450*2000",
                    "材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
                    "颜色": "黑色",
                    "数量": 3,
                    "单价": 10716.00,
                    "总价": 32148.00,
                    "备注": ""
                  },
                  {
                    "品名": "会客桌",
                    "规格": "2000*800*750",
                    "材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
                    "颜色": "黑色",
                    "数量": 6,
                    "单价": 4500.00,
                    "总价": 27000.00,
                    "备注": ""
                  },
                  {
                    "品名": "会客椅",
                    "规格": "常规",
                    "材质说明": "环保要求：甲醛释放量≤5mg/100g。\n2、基材：E0级\n3、木皮表面",
                    "颜色": "黑色",
                    "数量": 1,
                    "单价": 400.00,
                    "总价": 400.00,
                    "备注": ""
                  },
                  {
                    "品名": "中式隔断",
                    "规格": "定制2800*2100",
                    "材质说明": "采用行列式手法，风格、功能设计",
                    "颜色": "不锈钢包边",
                    "数量": 6,
                    "单价": 1200.00,
                    "总价": 7200.00,
                    "备注": ""
                  }
                ]
              }
            ]
          }
        } else {
          // 默认使用预算汇总表模板
          requestData.data = {
            sheets: [
              {
                name: '预算汇总表',
                title: '全屋智能家居方案A套餐预算汇总表',
                items: [
                  {
                    "序号": 1,
                    "品牌": "小米",
                    "区域": "全屋智能主控系统",
                    "系统说明": "1、AI智能语音、自定义设备各种场景（回家、离家、会客、就餐、休闲、阅读等模式），完美实现智能化体验。 2、智能品类包括：智能灯光、智能遮阳、智能空调，智能安防等； 3、最大优势及亮点 \"无缝接入米家APP、AI智能语音控制、轻成本、轻设计、轻方案、轻对接、轻落地、轻维护\"； 4、可以根据所需的智能开关与空调语音小助手进行DIY定制。",
                    "单位": "项",
                    "工程量": 1,
                    "预算价": 1928.00,
                    "单项预算合价": 1928.00
                  },
                  {
                    "序号": 2,
                    "品牌": "FSXRT",
                    "区域": "智能灯光",
                    "系统说明": "1、自定义色温：智能双色温（2700~6500K的灯具，可以根据需求DIY自定义设置色温参数； 2、控制方式：单灯控制、回路控制、互控、集成控制、远程控制等； 3、自定义氛围场景：娱乐、聚会、休闲、会客等灯光场景。",
                    "单位": "项",
                    "工程量": 1,
                    "预算价": 6759.00,
                    "单项预算合价": 6759.00
                  }
                ]
              }
            ]
          }
        }
      } else {
        // 其他文件类型使用默认数据结构
        requestData.data.sheets = [
          {
            name: 'Sheet1',
            headers: [['姓名', '年龄'], ['详细信息', '']],
            rows: [['张三', 25], ['李四', 30], ['王五', 28]]
          },
          {
            name: 'Sheet2',
            headers: [['产品', '价格', '数量']],
            rows: [['产品A', 100, 5], ['产品B', 200, 3], ['产品C', 150, 8]]
          }
        ]
      }

      const response = await axios.post(
        `http://localhost:8080/api/v1/export/${fileType}`,
        requestData,
        {
          responseType: 'blob',
          headers: {
            'Content-Type': 'application/json'
          }
        }
      )

      // 创建下载链接
      const url = window.URL.createObjectURL(new Blob([response.data]))
      const link = document.createElement('a')
      link.href = url
      link.setAttribute('download', `${formData.fileName}.${fileType == 'excel' ? 'xlsx' : fileType}`)
      document.body.appendChild(link)
      link.click()

      // 清理
      link.remove()
      window.URL.revokeObjectURL(url)

      setMessage(`成功导出${fileType.toUpperCase()}文件！`)
    } catch (error) {
      console.error('导出失败:', error)
      setMessage('导出失败，请检查服务是否正常运行')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>Office文件导出测试</h1>
        <p>测试Go语言Office文件导出服务</p>
      </header>

      <main className="app-main">
        <div className="form-section">
          <h2>导出配置</h2>
          <form className="export-form">
            <div className="form-group">
              <label htmlFor="title">报告标题:</label>
              <input
                type="text"
                id="title"
                name="title"
                value={formData.title}
                onChange={handleInputChange}
                required
              />
            </div>

            <div className="form-group">
              <label htmlFor="templateId">选择Excel模板:</label>
              <select
                id="templateId"
                name="templateId"
                value={formData.templateId}
                onChange={handleInputChange}
                required
              >
                {templates.map(template => (
                  <option key={template.id} value={template.id}>
                    {template.name} - {template.description}
                  </option>
                ))}
              </select>
            </div>

            <div className="form-group">
              <label htmlFor="fileName">文件名:</label>
              <input
                type="text"
                id="fileName"
                name="fileName"
                value={formData.fileName}
                onChange={handleInputChange}
                required
              />
            </div>
          </form>
        </div>

        <div className="export-section">
          <h2>开始导出</h2>
          <div className="export-buttons">
            <button
              onClick={() => exportFile('excel')}
              disabled={loading}
              className="export-btn excel"
            >
              {loading ? '导出中...' : '导出Excel'}
            </button>
            <button
              onClick={() => exportFile('word')}
              disabled={loading}
              className="export-btn word"
            >
              {loading ? '导出中...' : '导出Word'}
            </button>
            <button
              onClick={() => exportFile('pdf')}
              disabled={loading}
              className="export-btn pdf"
            >
              {loading ? '导出中...' : '导出PDF'}
            </button>
          </div>

          {message && (
            <div className={`message ${message.includes('成功') ? 'success' : 'error'}`}>
              {message}
            </div>
          )}
        </div>

        <div className="info-section">
          <h2>API信息</h2>
          <div className="api-info">
            <p><strong>服务地址:</strong> http://localhost:8080</p>
            <p><strong>导出接口:</strong> POST /api/v1/export/:type</p>
            <p><strong>支持类型:</strong> excel, word, pdf</p>
          </div>
        </div>
      </main>
    </div>
  )
}

export default App
