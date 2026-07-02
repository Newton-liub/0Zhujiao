# Pull Request

## 变更范围

- [ ] 功能或规则调整
- [ ] UI 调整
- [ ] 文档更新
- [ ] 构建或发布流程调整

## 数据安全检查

- [ ] 未提交真实 Excel、成绩表、学生名单或导出文件
- [ ] 未提交 token、密码、证书私钥或本地环境配置
- [ ] 如包含样例数据，已确认全部为虚构/脱敏数据

## 验证

- [ ] 已运行 `python -m compileall src`
- [ ] 如修改打包流程，已运行 `powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1`