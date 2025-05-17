# 将markdown文件转换为异步社区的书稿docx文件

## How to run the code

```shell
# Create a virtual environment in the current directory
python3 -m venv venv

# Activate the virtual environment
source venv/bin/activate

# Install the package in the virtual environment
pip install python-docx

# Now you can run your script
python converter.py ./book-vibe-coding-in-action/ch06/ch06.md

# Deactivate the virtual environment
deactivate
```

## 微调格式

- 代码清单标题格式微调：打开to-word-template.docx，用格式刷将代码清单的标题格式应用到转换后的DOCX文件中相应的代码块标题上。
- 代码清单代码体格式微调：选择“代码无行号”样式