# PPT Translator

基于 Ollama 的 PPT 翻译工具，支持中英互译，保持原始格式。

## 功能特点

- 支持中英互译
- 保持原始 PPT 格式和样式
- 智能字体大小调整
- 自动文本框大小调整
- 支持组合形状
- 支持多种翻译模型
- 图形界面操作

## 安装要求

- Python 3.8+
- PyQt6
- python-pptx
- requests

## 安装步骤

1. 克隆仓库：
```bash
git clone https://github.com/yourusername/ppt-translator.git
cd ppt-translator
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 图形界面

运行以下命令启动图形界面：
```bash
python ppt_translator_ui.py
```

### 命令行

也可以通过命令行使用：
```bash
python ppt_xml_translator.py --input input.pptx --output output.pptx --from-lang zh --to-lang en
```

参数说明：
- `--input`: 输入的 PPTX 文件路径
- `--output`: 输出的 PPTX 文件路径（可选）
- `--from-lang`: 源语言 (zh/en)
- `--to-lang`: 目标语言 (zh/en)
- `--model`: Ollama 模型名称
- `--host`: Ollama 服务地址

## 配置说明

1. 服务器配置
   - 默认服务器地址：`http://localhost:2342`
   - 需要先启动ollama服务，默认端口2342；或者获得ollama服务地址
   - 可在界面中修改或通过命令行参数指定

2. 支持的模型：
   - llama3:8b
   - qwen:7b
   - qwen:1.8b

## 注意事项

1. 确保有足够的磁盘空间用于临时文件
2. 建议在翻译前备份原始文件
3. 处理大型 PPT 文件时可能需要较长时间
4. 某些特殊格式可能需要手动调整

## 贡献指南

欢迎提交 Pull Request 或 Issue。

## 许可证

MIT License

## 作者

enzo X cursor