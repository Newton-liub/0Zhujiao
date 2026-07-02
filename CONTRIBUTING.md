# Contributing

感谢改进本项目。提交变更前请确认：

1. 不提交真实 Excel、成绩数据、学生名单或本地导出文件。
2. 新增功能保持数据流简单：输入 Excel → 标准化字段 → 合并排序 → 输出汇总表。
3. 不把 UI、Excel 解析、合并规则混在同一个模块里。
4. 修改后至少运行一次：

```powershell
python -m compileall src
```

5. 如果修改了打包逻辑，运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```