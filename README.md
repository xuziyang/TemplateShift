# TemplateShift - Excel数据转换工具

## 项目简介
TemplateShift是一个基于Electron和React的桌面应用程序，用于处理和转换Excel文件数据。它能够按照预定义的规则自动处理Excel文件，并生成符合特定格式要求的新Excel文件。

## 功能特点
- 支持拖拽上传Excel文件
- 自动识别和处理Excel文件内容
- 智能数据转换和整理：
  - 优先使用有效手机号，其次使用更多电话中的号码
  - 自动区分手机号和座机号
  - 智能处理地区信息（省份-城市）
  - 使用法定代表人作为客户名
  - 自动去重并保留最完整的记录
- 按地区自动排序
- 自动导出处理后的文件

## 技术栈
- Electron - 跨平台桌面应用框架
- React - 用户界面框架
- TypeScript - 开发语言
- Material-UI - UI组件库
- XLSX - Excel文件处理库
- Vite - 构建工具

## 开发环境配置
1. 克隆项目到本地：
```bash
git clone [项目地址]
cd TemplateShift
```

2. 安装依赖：
```bash
npm install
```

3. 启动开发服务器：
```bash
npm run dev
```

## 使用说明
1. 启动应用程序
2. 通过以下两种方式之一上传Excel文件：
   - 将Excel文件拖拽到指定区域
   - 点击"选择文件"按钮选择Excel文件
3. 程序将自动处理文件并导出转换后的新Excel文件

## 数据处理规则
- 输入文件要求：
  - 支持.xlsx和.xls格式
  - 第二行为表头行
  - 从第三行开始为数据内容
- 数据转换规则：
  - 电话号码优先级：有效手机号 > 座机号 > 更多电话中的号码
  - 地区信息：省份和城市自动合并为"省份-城市"格式
  - 客户名：优先使用法定代表人
  - 自动去重：基于电话号码去重，保留最完整的记录

## 输出文件格式
- 文件名：原文件名_processed.xlsx
- 文件结构：
  - 1-12行：空行
  - 第13行：表头（客户名、手机、微信、来源、地区、备注）
  - 第14行起：转换后的数据

## 注意事项
- 确保Excel文件格式正确
- 表头行必须位于第二行
- 数据必须从第三行开始
- 建议在处理大文件时保持耐心等待