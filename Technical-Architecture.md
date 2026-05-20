# ExcelDiff - 技术架构文档

## 1. 架构设计

### 系统架构图

```mermaid
graph LR
    A[用户界面层] --> B[业务逻辑层]
    B --> C[数据处理层]
    C --> D[文件解析层]
    
    A[用户界面层] -->|上传文件| B
    B -->|比较结果| A
    D -->|Excel数据| C
```

### 层级说明

```
┌─────────────────────────────────────┐
│      用户界面层 (UI Layer)           │
│  - 文件上传组件                      │
│  - 差异展示组件                      │
│  - 统计面板组件                      │
│  - 导出功能组件                      │
└─────────────────────────────────────┘
                ↑
┌─────────────────────────────────────┐
│      业务逻辑层 (Business Layer)     │
│  - Excel文件解析                     │
│  - 数据比较算法                      │
│  - 差异分类逻辑                      │
│  - 报告生成逻辑                      │
└─────────────────────────────────────┘
                ↑
┌─────────────────────────────────────┐
│      数据处理层 (Data Layer)         │
│  - Excel数据转换为JSON               │
│  - 单元格级差异检测                  │
│  - 差异数据结构化                    │
└─────────────────────────────────────┘
```

---

## 2. 技术栈

### 前端框架
- **React 18** - UI框架
- **Vite** - 构建工具和开发服务器
- **TailwindCSS** - 样式框架

### 核心库
- **xlsx (SheetJS)** - Excel文件解析和生成
- **lucide-react** - 图标库

### 开发工具
- npm - 包管理
- ESLint - 代码检查
- Prettier - 代码格式化

---

## 3. 项目结构

```
excel-compare/
├── public/
│   └── favicon.svg
├── src/
│   ├── components/
│   │   ├── FileUploader.jsx      # 文件上传组件
│   │   ├── DiffViewer.jsx       # 差异展示组件
│   │   ├── StatsPanel.jsx        # 统计面板组件
│   │   └── ExportButton.jsx      # 导出按钮组件
│   ├── utils/
│   │   ├── excelParser.js        # Excel解析工具
│   │   ├── diffEngine.js         # 差异比较引擎
│   │   └── reportGenerator.js    # 报告生成工具
│   ├── App.jsx                   # 主应用组件
│   ├── App.css                   # 全局样式
│   ├── index.css                 # 入口样式
│   └── main.jsx                  # 入口文件
├── index.html
├── package.json
├── vite.config.js
├── tailwind.config.js
└── postcss.config.js
```

---

## 4. 核心模块设计

### 4.1 Excel解析模块 (excelParser.js)

**功能：** 解析Excel文件为JSON格式

**核心函数：**
```javascript
parseExcel(file) 
  → 解析上传的Excel文件
  → 返回 { sheets: [{ name, data: [[...], ...] }] }

getSheetNames(workbook)
  → 获取所有工作表名称

sheetToJson(sheet)
  → 将单个工作表转换为二维数组
```

**依赖：** xlsx库

---

### 4.2 差异比较引擎 (diffEngine.js)

**功能：** 比较两个Excel数据集，找出差异

**核心函数：**
```javascript
compareData(data1, data2)
  → 输入两个二维数组
  → 返回差异数组

diffTypes:
  - 'added': 在data2中新增的行/单元格
  - 'deleted': 在data1中存在但在data2中不存在的
  - 'modified': 值发生变化的
```

**比较策略：**
1. 先按行列索引对齐
2. 逐单元格比较值
3. 分类差异类型
4. 生成差异报告

---

### 4.3 报告生成器 (reportGenerator.js)

**功能：** 生成可导出的HTML差异报告

**核心函数：**
```javascript
generateHTMLReport(diffResults, fileName1, fileName2)
  → 生成包含样式的HTML字符串
  → 保留颜色高亮
  → 支持下载
```

---

## 5. 组件设计

### 5.1 FileUploader 组件

**功能：** 处理文件拖拽和上传

**Props：**
```javascript
{
  onFileSelect: (file) => void,  // 文件选择回调
  file: File | null,             // 已选择的文件
  label: string                  // 上传区域标签
}
```

**状态：**
- `isDragging`: 是否正在拖拽
- `file`: 已选择的文件对象

**事件处理：**
- `onDragEnter`: 开始拖拽
- `onDragLeave`: 拖拽离开
- `onDrop`: 文件放下
- `onChange`: 文件选择变化

---

### 5.2 DiffViewer 组件

**功能：** 展示差异结果

**Props：**
```javascript
{
  diffResults: {
    added: [...],
    deleted: [...],
    modified: [...],
    stats: { total, added, deleted, modified }
  },
  originalData: [[...]],  // 原始数据
  modifiedData: [[...]]   // 修改后数据
}
```

**展示内容：**
- 统计摘要（饼图或数字卡片）
- 差异列表（可滚动）
- 差异详情（单元格位置、旧值、新值）

---

### 5.3 StatsPanel 组件

**功能：** 显示比较统计信息

**Props：**
```javascript
{
  stats: {
    totalCells: number,
    added: number,
    deleted: number,
    modified: number,
    identical: number
  }
}
```

**展示形式：**
- 数字卡片 + 图标
- 颜色编码（绿/红/黄/灰）
- 百分比显示

---

## 6. 数据模型

### 6.1 Excel数据结构

```typescript
interface ExcelData {
  sheets: Sheet[];
}

interface Sheet {
  name: string;
  data: Cell[][];  // 二维数组
}

interface Cell {
  value: string | number | null;
  row: number;
  col: number;
}
```

### 6.2 差异数据结构

```typescript
interface DiffResult {
  type: 'added' | 'deleted' | 'modified';
  row: number;
  col: number;
  oldValue?: any;
  newValue?: any;
  position: string;  // 例如: "Sheet1!A1"
}

interface ComparisonResult {
  diffs: DiffResult[];
  stats: {
    total: number;
    added: number;
    deleted: number;
    modified: number;
  };
}
```

---

## 7. 样式设计

### TailwindCSS配置

```javascript
// tailwind.config.js
module.exports = {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        primary: '#6366F1',
        success: '#22C55E',
        danger: '#EF4444',
        warning: '#F59E0B',
        dark: {
          bg: '#0F172A',
          card: '#1E293B',
          border: '#334155'
        }
      },
      fontFamily: {
        heading: ['Poppins', 'sans-serif'],
        body: ['Inter', 'sans-serif']
      }
    }
  }
}
```

### 动画效果

```css
/* 拖拽上传动画 */
.upload-zone {
  transition: all 0.3s ease;
}
.upload-zone.dragging {
  border-color: #6366F1;
  transform: scale(1.02);
}

/* 数字计数动画 */
.stat-number {
  animation: countUp 0.6s ease-out;
}

@keyframes countUp {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}
```

---

## 8. 性能优化

### 文件处理优化
- 使用FileReader API异步读取
- 大文件分块处理（可选）
- 避免阻塞主线程

### 渲染优化
- 使用React.memo优化组件重渲染
- 虚拟列表处理大量差异项（如果需要）
- 懒加载非关键组件

### 用户体验优化
- 加载状态指示
- 错误提示友好
- 支持拖拽和点击两种上传方式

---

## 9. 浏览器兼容性

| 浏览器 | 最低版本 | 推荐版本 |
|--------|----------|----------|
| Chrome | 80+ | 最新版 |
| Firefox | 75+ | 最新版 |
| Safari | 13+ | 最新版 |
| Edge | 80+ | 最新版 |

---

## 10. 安全性考虑

- 所有文件处理在浏览器本地完成
- 不上传文件到任何服务器
- 不使用cookies或本地存储存储文件数据
- 文件数据仅存在于内存中
