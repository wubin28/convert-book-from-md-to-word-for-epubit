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

- 代码清单“```markdown”标题格式微调：打开to-word-template.docx，用格式刷将代码清单的标题格式应用到转换后的DOCX文件中相应的代码块标题上。
- 代码清单“```markdown”代码体格式微调：选择“代码无行号”样式
- "```shell"格式微调：选择“代码无行号”样式

## 初始提示词

```markdown
我上传了3个文件，其中 ch04-from.md 和 ch04-to-original.docx 是两个格式不同但内容相同的文件。我需要一个名为 converter 的 Python 程序来分析这两个文件的格式差异。当运行 "python3 converter ch04-from.md" 时，程序应执行以下操作：读取 ch04-from.md 的内容，复制 ch04-to-template.docx 文件并重命名为 ch04-to.docx，然后将 ch04-from.md 中的内容按照 ch04-to-original.docx 的格式写入 ch04-to.docx 中。转换完成后，用 Word 打开 ch04-to.docx 时应与 ch04-to-original.docx 的效果完全一致。由于没有上传 markdown 文件中的图片，转换后的 ch04-to.docx 可以不包含图片，但所有文字内容和格式必须与原文件保持一致，不能增减。如遇到"【避坑指南】"这样的特殊格式无法确定如何转换，请告知并尽力保留这些内容。
```