import os
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
from docx import Document

def select_directory():
    """使用 tkinter 唤起沙盒目录选择"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title="请选择要处理的 Word 文档根目录")
    return Path(folder_path) if folder_path else None

def process_word_documents():
    input_dir = select_directory()
    if not input_dir or not input_dir.exists():
        print("未选择有效目录，程序退出。")
        return

    # 定义并创建输出目录（在输入目录同级，防止无限套娃遍历）
    output_dir = input_dir.parent / f"{input_dir.name}_Cleaned_Output"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"✅ 输入目录: {input_dir}")
    print(f"✅ 输出目录: {output_dir}\n")

    audit_records = []

    # 递归遍历所有 .docx 文件
    for file_path in input_dir.rglob("*.docx"):
        # 1. 排除临时文件
        if file_path.name.startswith("~$"):
            continue
        
        # 2. 严防无限套娃：跳过已经在输出目录中的文件
        if output_dir in file_path.parents:
            continue

        rel_path = file_path.relative_to(input_dir)
        out_file_path = output_dir / rel_path
        
        # 3. 原样复刻多级子文件夹结构
        out_file_path.parent.mkdir(parents=True, exist_ok=True)

        status = "成功"
        error_msg = ""
        fixed_nodes = 0

        try:
            doc = Document(file_path)
            
            # 遍历文档中的所有段落
            for p in doc.element.xpath('//w:p'):
                is_hyperlink_field = False
                
                # 严格遵守 API 规范：获取所有子孙节点
                for element in p.xpath('.//*'):
                    # 动态获取标签名
                    tag_name = element.tag.split('}')[-1]
                    
                    # 追踪域代码状态 (Track 2 保护: <w:instrText> HYPERLINK)
                    if tag_name == 'fldChar':
                        # 降维打击：使用 local-name() 替代 @w:fldCharType，彻底根除命名空间报错
                        fld_types = element.xpath('@*[local-name()="fldCharType"]')
                        if fld_types:
                            if fld_types[0] == 'begin':
                                is_hyperlink_field = False
                            elif fld_types[0] == 'end':
                                is_hyperlink_field = False
                                
                    elif tag_name == 'instrText':
                        if element.text and 'HYPERLINK' in element.text.upper():
                            is_hyperlink_field = True
                            
                    # 处理文本块 <w:r>
                    elif tag_name == 'r':
                        # Track 1 保护: 使用 local-name() 寻找祖先节点，替代 ancestor::w:hyperlink
                        in_w_hyperlink = bool(element.xpath('ancestor::*[local-name()="hyperlink"]'))
                        
                        # 如果命中了上述两种真实链接的任何一种，直接放行
                        if in_w_hyperlink or is_hyperlink_field:
                            continue
                            
                        # === 暴力的样式清洗 (无差别斩首) ===
                        # 使用 local-name() 安全获取 rPr
                        rPrs = element.xpath('./*[local-name()="rPr"]')
                        if rPrs:
                            rPr = rPrs[0]
                            # 使用 local-name() 无条件抓取 rStyle, color, u 标签，避开 w: 前缀
                            tags_to_remove = (
                                rPr.xpath('./*[local-name()="rStyle"]') + 
                                rPr.xpath('./*[local-name()="color"]') + 
                                rPr.xpath('./*[local-name()="u"]')
                            )
                            
                            if tags_to_remove:
                                for tag in tags_to_remove:
                                    rPr.remove(tag)
                                fixed_nodes += 1

            # 保存清洗后的文档
            doc.save(out_file_path)
            print(f"处理完成 [{fixed_nodes} 处修复]: {file_path.name}")

        except Exception as e:
            status = "失败"
            error_msg = str(e)
            print(f"❌ 处理异常: {file_path.name} | 报错: {error_msg}")

        # 记录审计数据
        audit_records.append({
            "文件名": file_path.name,
            "相对路径": str(rel_path),
            "处理状态": status,
            "修复节点数": fixed_nodes,
            "错误信息": error_msg
        })

    # 4. 生成 Excel 审计报告
    if audit_records:
        df = pd.DataFrame(audit_records)
        report_path = output_dir / "伪链接清理审计报告.xlsx"
        df.to_excel(report_path, index=False)
        print(f"\n🎉 批量任务执行完毕！")
        print(f"📊 审计报告已生成: {report_path}")
    else:
        print("\n未找到任何可处理的 .docx 文件。")

if __name__ == "__main__":
    process_word_documents()