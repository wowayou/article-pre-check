# SEO Super Engine V6.5 - 维护与架构手册 (Maintenance Manual)

这份文档旨在为后续的开发人员、架构师以及运维人员提供代码库的深层技术解析与维护指南。

## 🏗️ 架构概览 (Architecture Overview)

脚本严格遵循面向对象设计 (OOP) 与单一职责原则 (SRP)，分为两层核心架构：

### 1. 调度与界面层 (`SEOWorkflowManagerV65`)
- **生命周期控制**: 使用 `tkinter` 构建原生 GUI，负责挂载用户配置。处理 `root.withdraw()` 与 `temp_root` 的切换以防止对话框 `RuntimeError` 崩溃。
- **并发引擎**: `execute_all` 方法利用 `concurrent.futures.ThreadPoolExecutor`，将针对 `os.walk` 提取的每一个独立文件下发给线程池。
- **沙箱机制**: 为避免污染源数据，所有任务均基于相对路径 (`rel_path`) 被映射到父级目录的 `_Cleaned_Output` 沙箱中。

### 2. 单例处理引擎 (`SEOSuperEngineV65`)
负责处理单一 `Document` 的一切 DOM 操作。完全无锁，保证了上层多线程调用的绝对线程安全。
- 采用 **生成器表达式 (Generator Expressions)** 代替列表推导式提取文本，极大降低了超大 Word 文档解析时的内存峰值。
- 所有异常流（如特定格式的损坏节点）均由 `try...except Exception as e:` 捕获并通过标准 `logging` 模块降级处理，绝不引发整个批处理中断。

---

## 🧠 核心算法揭秘 (Core Algorithms)

### 1. 标题层级修复 (Level Tracker)
**问题背景**: 原生 Word 很容易出现 `H1 -> H3 -> H3` 的跳级（即跳过了 H2）。如果采用简单的绝对字典映射，当后续恢复到正常的 `H2` 时，极易产生将合法 `H2` 错误提升至 `H1` 的惨剧。
**算法实现**: `_fix_heading_hierarchy`
- 采用 **相对偏移量 (`current_offset`)** 追踪实际层级与应有层级的差值。
- **降级触发**: `potential_lv = actual_lv - current_offset`。当 `potential_lv > last_fixed_lv + 1`，则判定为跳级，增大 `current_offset`。
- **自愈机制**: 如果 `actual_lv <= last_fixed_lv` (遇到了上层合法父标题)，且不构成对前一个合法节点的跳级，则动态将 `current_offset` 收缩或清零，从而保护原有的大纲框架。

### 2. 伪链接的斩首式清洗 (Robust Link Cleanup)
**问题背景**: 基础的 `python-docx` 只能清理 XML 结构树上的 `<w:hyperlink>` 节点，但 Word 依然会渲染蓝色下划线。这是因为 `<w:rPr>` 中残留了样式代码。如果使用 `w:val="Hyperlink"` 去匹配，经常因各种环境下的 XML 命名空间问题而失败。
**算法实现**: `_strip_run_hyperlink_style_robust`
- 利用 `lxml` 的 XPath 能力，直接降维打击：使用 `local-name()` 绕过前缀命名空间。
- `.xpath('./*[local-name()="rPr"]')` 无差别提取属性节点，暴力移除底层的 `rStyle`, `color`, `u`，从根源上阻断样式的残留渲染。

---

## 🛠️ 故障排查 (Troubleshooting)

### Q1: `tkinter` 抛出 `Too early to create variable`
- **原因**: 通常是在某个地方过早地调用了主窗口的 `destroy()`，导致后续创建弹窗（如 `messagebox`）时找不到默认的 root 父级上下文。
- **对策**: 脚本中已通过 `temp_root = tk.Tk(); temp_root.withdraw()` 手动挂载隐形父级修复。不要轻易去掉这些上下文挂载代码。

### Q2: 导出的 Excel 数据串行或错乱
- **原因**: 之前版本若使用共享 List 的 `append` 在多线程下是不安全的。
- **对策**: 现在的架构是在 `ThreadPoolExecutor` 中使用 `as_completed(tasks)` 收集全部独立的 `Future` 结果后，再单线程统合到 `self.results` 中，不要在 Task 内部直接操作 `self.results`。

### Q3: `skip_rules.json` 规则未生效
- **原因**: 字符串大小写敏感匹配问题。
- **对策**: 请确保代码中 `kw.lower() in file_name_lower` 的对比逻辑不被更改，配置 JSON 中的关键词对大小写不敏感。