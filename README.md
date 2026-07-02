# Excel 成绩汇总工具

这个工具用于把一个或多个考试 Excel 汇总成登分表。

## 输出内容

最终汇总表只保留：

- 学号/工号
- 学生姓名
- 班级
- 每个源 Excel 对应的一列成绩

例如导入：

- `单片机阶段测试1.xlsx`
- `单片机阶段测试2.xlsx`

输出列会类似：

- 学号/工号
- 学生姓名
- 班级
- 阶段测试1
- 阶段测试2

## 排序规则

- 自动识别最多出现的年级作为主年级。
- 与主年级不同的班级视为重修，放在最上面。
- 其他学生按班级从小到大排列。
- 同班级内按学号/工号从小到大排列。

## 下载 Windows 版

不想运行源码的用户，可以在 GitHub Releases 下载 `ScoreTool.exe`。

## 直接运行源码

```powershell
python -m pip install -r requirements.txt
$env:PYTHONPATH = "$PWD\src"
python -m score_tool.app
```

## 打包 exe

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

打包后 exe 位置：

```text
dist\ScoreTool\ScoreTool.exe
```

## 使用步骤

1. 双击打开 `ScoreTool.exe`。
2. 点击 `添加 Excel`，选择一个或多个成绩表。
3. 检查每个文件右侧的成绩列名。
4. 如果列名不合适，选中文件后点击 `修改成绩列名`。
5. 点击 `选择保存位置`。
6. 点击 `预检`，确认表头、人数、警告信息。
7. 点击 `生成汇总表`。
8. 打开输出文件检查结果。

## 数据安全

工具只读取并导出四类数据：

- 学号/工号
- 学生姓名
- 班级
- 总分

不会把 IP、提交时间、答题明细、题目分数等列写入汇总表。