# MindManager COM Interface Generator

这是一个用于生成 MindManager COM 接口 Python 包装器的工具。该工具可以自动分析 COM 对象的接口，并生成对应的 Python 类，使得在 Python 中操作 MindManager 变得更加简单。

## 功能特点

- 自动分析 COM 对象的接口和属性
- 生成类型安全的 Python 包装类
- 支持所有 COM 对象的属性和方法
- 自动处理 COM 对象的生命周期
- 提供详细的类型信息和文档字符串

## 项目结构

```
mindmanager-pluggin/
├── com_interface.py      # COM 接口生成器主程序
├── generated/            # 生成的接口代码目录
│   ├── __init__.py      # 包初始化文件
│   └── *.py            # 生成的接口类文件
└── README.md            # 项目说明文档
```

## 环境要求

- Python 3.6+
- pywin32
- Windows 操作系统

## 安装

1. 克隆项目：
```bash
git clone https://github.com/111hgx/comobj2py.git
cd comobj2py
```

2. 安装依赖：
```bash
uv add pywin32
```

## 使用方法

1. 运行接口生成器：
```bash
python comobj2py
```

2. 输入 COM 对象名称，例如：
- `MindManager.Application` - 用于 MindManager
- `Word.Application` - 用于 Word
- 等等

3. 生成的代码将保存在 `generated` 目录中

## 示例代码

### 使用 MindManager 接口

```python
import win32com.client
import pythoncom
from generated import ActiveDocument, Documents, Options


def print_topic(topic, level=0):
    """递归打印主题及其子主题"""
    indent = "  " * level
    print(f"{indent}主题: {topic.Text}")

    # 获取子主题
    subtopics = topic.SubTopics
    if subtopics:
        print(f"{indent}子主题数量: {subtopics.Count}")
        for i in range(subtopics.Count):
            subtopic = subtopics.Item(i + 1)  # COM索引从1开始
            print_topic(subtopic, level + 1)


def main():
    try:
        # 初始化COM
        pythoncom.CoInitialize()

        # 创建MindManager应用程序实例
        app = win32com.client.Dispatch("MindManager.Application")

        # 获取活动文档
        active_doc = ActiveDocument()

        # 获取文档集合
        docs = Documents()

        # 获取选项
        options = Options()

        # 打印一些基本信息
        print(f"用户名: {options.UserName}")
        print(f"用户邮箱: {options.UserEmail}")
        print(f"文档数量: {docs.Count}")

        # 创建一个新文档
        # new_doc = docs.Add()
        # new_doc.Activate()  # 激活新创建的文档,打开mindmanager窗口
        # print(f"已创建新文档")

        # 打印当前选项值
        print(f"\n当前选项:")
        print(f"显示工作簿标签: {options.ShowWorkbookTabs}")
        print(f"在工作簿标签中显示文件名: {options.ShowFilenamesInWorkbookTabs}")

        # 获取当前打开的文档
        doc = ActiveDocument()
        if not doc:
            print("没有打开的导图文档")
            return

        print(f"\n当前导图: {doc.Name}")

        # 获取中心主题
        central_topic = doc.CentralTopic
        if central_topic:
            print("\n导图结构:")
            print_topic(central_topic)
        else:
            print("未找到中心主题")

        print("\n示例运行完成")

    except Exception as e:
        print(f"发生错误: {str(e)}")
    finally:
        # 清理COM
        pythoncom.CoUninitialize()
        # 清理COM对象
        try:
            app.Quit()
        except:
            pass


if __name__ == "__main__":
    main()

```

## 注意事项

1. 确保在运行程序前已安装所需的 COM 对象（如 MindManager）
2. 某些 COM 对象可能需要管理员权限才能访问
3. 生成的代码可能需要根据具体使用场景进行调整
4. 建议在使用前先测试生成的接口是否满足需求

## 常见问题

1. **Q: 为什么某些属性或方法无法访问？**  
   A: 这可能是因为 COM 对象的权限限制或接口未完全实现。请检查 COM 对象的文档。

2. **Q: 如何处理 COM 对象的错误？**  
   A: 使用 try-except 块捕获可能的异常，并检查错误信息。

3. **Q: 生成的代码是否需要修改？**  
   A: 生成的代码通常可以直接使用，但可能需要根据具体需求进行调整。

## 贡献

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目。

## 许可证

```text
MIT License

Copyright (c) [2025] [111hgx]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```