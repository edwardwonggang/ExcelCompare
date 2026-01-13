# ExcelCompare - Excel文件智能比对工具

## 功能说明

本工具用于在Git提交前自动检查Excel文件，防止覆盖他人修改，确保协作安全。

### 主要功能

1. **版本一致性检查**：检查本地修改的Excel文件是否基于最新的远程版本
2. **修订记录检查**：检查Excel文件中的"修改记录"sheet页是否有新的修订记录
3. **批量检查优化**：支持多线程检查，确保100个文件也能快速完成
4. **提交拦截**：检查失败时自动拦截Git提交

## 安装步骤

### 1. 安装依赖

```bash
pip install openpyxl pandas
```

### 2. 安装Git钩子

```bash
python install_hooks.py
```

### 3. 配置检查规则

编辑 `config.json` 文件，配置检查规则：

```json
{
  "sheet_name": "修改记录",
  "check_columns": ["修订人", "修订时间", "修订内容"],
  "max_threads": 10,
  "timeout": 30
}
```

## 使用说明

### 自动检查

安装钩子后，每次执行 `git commit` 时会自动检查暂存区中的Excel文件。

### 手动检查

```bash
python excel_checker.py
```

### 检查所有Excel文件

```bash
python excel_checker.py --all
```

## 工作原理

1. **pre-commit钩子**：在提交前触发检查
2. **版本比对**：比较本地文件与远程最新版本
3. **修订记录分析**：解析"修改记录"sheet页，检查是否有新记录
4. **多线程处理**：使用线程池并行检查多个文件

## 注意事项

1. 确保在提交前先执行 `git pull` 获取最新代码
2. "修改记录"sheet页必须包含指定的列名
3. 检查失败时，请根据提示更新文件或添加修订记录

## 卸载

```bash
python uninstall_hooks.py
```
