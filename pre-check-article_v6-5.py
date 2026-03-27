import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Optional, Tuple, List
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
import logging

# ==========================================
# 核心数据类 (Dataclass): 规范输出数据结构
# ==========================================
@dataclass
class AuditResult:
    project_folder: str
    file_name_link: str
    domain_status: str
    h1_status: str
    heading_status: str
    missing_images: str
    company_info: str
    tdk_content: str
    tdk_advice: str
    all_links: str
    cleaned_links_count: str
    logs: str

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ==========================================
# 核心处理引擎
# ==========================================
class SEOSuperEngineV65:
    """企业级 Word 文档 SEO 自动化审查与修复引擎"""
    
    def __init__(self, file_path: Path, target_domain: Optional[str], folder_files: List[str], config: dict):
        self.file_path = file_path
        self.doc = Document(self.file_path)
        self.folder_files = folder_files
        self.target_domain = self._normalize_domain(target_domain)
        self.config = config  # 接收 GUI 配置面板的参数
        
        # 状态记录器
        self.changes = []
        self.missing_images = []
        self.links_removed_count = 0
        self.all_links_found = set()

    def _normalize_domain(self, domain_str: Optional[str]) -> Optional[str]:
        """清洗提取主域名"""
        if not domain_str or str(domain_str).strip() == 'nan': return None
        d = str(domain_str).strip().lower()
        d = re.sub(r'^https?://', '', d)
        d = re.sub(r'^www\.', '', d)
        return d.split('/')[0]

    def _is_external(self, url: str) -> bool:
        """判定是否为非本站外链"""
        if not self.target_domain or not url: return False
        url_lower = url.lower()
        if url_lower.startswith(('#', 'mailto:')): return False
        url_clean = re.sub(r'^https?://', '', url_lower)
        url_clean = re.sub(r'^www\.', '', url_clean)
        return self.target_domain not in url_clean

    def _strip_run_hyperlink_style(self, run_el):
        """卸载文本节点上的超链接样式 (颜色、下划线)"""
        rPr = run_el.find(qn('w:rPr'))
        if rPr is not None:
            rStyle = rPr.find(qn('w:rStyle'))
            if rStyle is not None and rStyle.get(qn('w:val')) == 'Hyperlink':
                rPr.remove(rStyle)
            color = rPr.find(qn('w:color'))
            if color is not None: rPr.remove(color)
            u = rPr.find(qn('w:u'))
            if u is not None: rPr.remove(u)

    def _clean_links(self, apply_fix: bool):
        """外链清理核心逻辑 (支持标准节点与域代码)"""
        rels = self.doc.part.rels
        links_to_remove = []
        field_codes_to_remove = []

        # 1. 扫描标准节点与纯文本
        for p in self.doc.paragraphs:
            for hl in p._element.xpath('.//w:hyperlink'):
                rId = hl.get(qn('r:id'))
                if rId in rels and rels[rId]._target:
                    url = rels[rId]._target
                    self.all_links_found.add(url)
                    if self._is_external(url): links_to_remove.append((hl, url))
                    
        # 2. 扫描域代码 (Field Codes)
        for instr in self.doc.element.xpath('.//w:instrText'):
            if instr.text and "HYPERLINK" in instr.text:
                url_match = re.search(r'"(https?://[^"]+)"', instr.text)
                if url_match:
                    url = url_match.group(1)
                    self.all_links_found.add(url)
                    if self._is_external(url): field_codes_to_remove.append((instr, url))

        # 3. 扫描纯文本 (仅提取，无法直接安全删除)
        full_text = "\n".join([p.text for p in self.doc.paragraphs if p.text])
        for url in re.findall(r'(https?://[^\s<>"]+|www\.[^\s<>"]+)', full_text):
            self.all_links_found.add(url)

        # 仅在配置允许且有目标域名时执行清理
        if self.config['clean_links'] and self.target_domain:
            for hl, url in links_to_remove:
                if apply_fix:
                    try:
                        parent = hl.getparent()
                        if parent is not None:
                            for child in list(hl):
                                if child.tag.endswith('}r'): self._strip_run_hyperlink_style(child)
                                hl.addprevious(child)
                            parent.remove(hl)
                            self.links_removed_count += 1
                    except: pass
                else: self.changes.append(f"待清理标准外链: {url}")
                
            for instr, url in field_codes_to_remove:
                if apply_fix:
                    try:
                        instr.text = instr.text.replace('HYPERLINK', 'QUOTE')
                        parent_r = instr.getparent()
                        if parent_r is not None:
                            curr = parent_r.getnext()
                            while curr is not None:
                                if curr.xpath('.//w:fldChar[@w:fldCharType="end"]'): break
                                if curr.tag.endswith('}r'): self._strip_run_hyperlink_style(curr)
                                curr = curr.getnext()
                        self.links_removed_count += 1
                    except: pass
                else: self.changes.append(f"待清理域代码外链: {url}")

    def _fix_images_and_headings(self, apply_fix: bool) -> str:
        """处理图片标注与层级跳级"""
        img_re = re.compile(r'(img\.)(.*?)\.(jpg|jpeg|png|bmp|gif|webp)', re.I)
        headings = []
        
        for p in self.doc.paragraphs:
            has_tag = "img." in p.text.lower()
            # 穿透 XML 查找真实图片对象
            has_obj = len(p._element.xpath('.//w:drawing | .//w:pict')) > 0
            
            # 如果图片(标注或对象)被错误地设置为了标题格式
            if (has_tag or has_obj) and p.style.name.startswith(('Heading', '标题')):
                if apply_fix: p.style = 'Normal'
                self.changes.append("样式修复: 图片/标注行降级为正文")
            
            # 处理图片标注文本
            if has_tag:
                matches = list(img_re.finditer(p.text))
                new_text = p.text
                for m in matches:
                    clean_fname = re.sub(r'[\\/:*?"<>|]', '', m.group(2))
                    orig_ext = m.group(3).lower()
                    
                    # 逻辑分流：是否强制 .webp
                    target_ext = "webp" if self.config['force_webp'] else orig_ext
                    fixed_tag = f"{m.group(1)}{clean_fname}.{target_ext}"
                    
                    # 检查文件夹中是否存在该图
                    if f"{clean_fname}.{target_ext}".lower() not in self.folder_files:
                        self.missing_images.append(f"{clean_fname}.{target_ext}")
                        
                    if m.group(0) != fixed_tag:
                        if apply_fix:
                            new_text = new_text.replace(m.group(0), fixed_tag)
                            self.changes.append(f"标注修正: {fixed_tag}")
                if apply_fix: p.text = new_text

            # 收集正常标题用于跳级检查
            if p.style.name.startswith(('Heading', '标题')) and not ("img." in p.text.lower() or has_obj):
                m_level = re.search(r'\d', p.style.name)
                if m_level: headings.append((p, int(m_level.group())))

        # 修复跳级
        h_status = "正常"
        last_lv = 0
        for p, curr_lv in headings:
            if last_lv > 0 and curr_lv > last_lv + 1:
                new_lv = last_lv + 1
                if apply_fix:
                    pref = "Heading " if "Heading" in p.style.name else "标题 "
                    try: p.style = f"{pref}{new_lv}"
                    except: pass
                h_status = f"层级跳级(H{curr_lv}->H{new_lv})"
                self.changes.append(h_status)
            last_lv = curr_lv
            
        return h_status

    def _check_h1_uniqueness(self) -> str:
        """检查 H1 唯一性"""
        h1_count = 0
        for p in self.doc.paragraphs:
            if p.style.name.startswith(('Heading 1', '标题 1')):
                # 排除图片对象的干扰
                if not ("img." in p.text.lower() or len(p._element.xpath('.//w:drawing | .//w:pict')) > 0):
                    h1_count += 1
        
        if h1_count == 0: return "❌ 缺失 H1"
        elif h1_count == 1: return "正常 (1个)"
        else: return f"❌ 异常 ({h1_count}个 H1)"

    def _extract_company_info(self) -> str:
        """提取公司描述信息"""
        for p in self.doc.paragraphs:
            full_txt = "".join([node.text for node in p._element.xpath('.//w:t') if node.text])
            if "co., ltd" in full_txt.lower():
                return p.text.strip()
        return "未发现"

    def process(self, apply_fix=False) -> Tuple[str, str, str, str, str, str]:
        """执行单一文件的全面审计与修复"""
        self._clean_links(apply_fix)
        h_status = self._fix_images_and_headings(apply_fix)
        co_info = self._extract_company_info()
        h1_status = self._check_h1_uniqueness() if self.config['enable_seo_check'] else "未开启检查"

        if apply_fix and (self.changes or self.links_removed_count > 0):
            try: self.doc.save(self.file_path)
            except Exception as e: logging.error(f"保存失败 {self.file_path.name}: {e}")
            
        logs = "; ".join(set(self.changes)) if self.changes else "正常"
        all_links_str = "\n".join(list(self.all_links_found))
        miss_imgs_str = ", ".join(set(self.missing_images)) if self.missing_images else "无"
        
        return h1_status, h_status, co_info, all_links_str, miss_imgs_str, logs

# ==========================================
# 工作流与 GUI 管理器
# ==========================================
class SEOWorkflowManagerV65:
    def __init__(self):
        self.domain_map = {}
        self.results = []
        self.root = tk.Tk()
        self.root.title("SEO Super Engine V6.5 - 控制面板")
        self.root.geometry("480x420")
        self.root.eval('tk::PlaceWindow . center')
        
        # 配置变量
        self.var_clean_links = tk.BooleanVar(value=True)
        self.var_force_webp = tk.BooleanVar(value=False)
        self.var_seo_check = tk.BooleanVar(value=True)
        self.run_status = False

    def _build_gui(self):
        """构建专业级配置界面"""
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="⚙️ 核心任务配置", font=("微软雅黑", 12, "bold")).pack(anchor=tk.W, pady=(0, 15))

        # 选框 1: 外链清理
        cb1 = ttk.Checkbutton(frame, text="开启【自动清理非本站外链】", variable=self.var_clean_links)
        cb1.pack(anchor=tk.W, pady=5)
        ttk.Label(frame, text="  (包含标准链接与隐藏域代码链接)", foreground="gray").pack(anchor=tk.W, pady=(0, 10))

        # 选框 2: Webp 强制统一
        cb2 = ttk.Checkbutton(frame, text="开启【强制统一图片标注后缀为 .webp】", variable=self.var_force_webp)
        cb2.pack(anchor=tk.W, pady=5)
        ttk.Label(frame, text="  (不勾选则严格校验原后缀如 .jpg 是否与文件夹一致)", foreground="gray").pack(anchor=tk.W, pady=(0, 10))

        # 选框 3: 深度 SEO 审查
        cb3 = ttk.Checkbutton(frame, text="开启【深度 SEO 内容审查】", variable=self.var_seo_check)
        cb3.pack(anchor=tk.W, pady=5)
        ttk.Label(frame, text="  (包含 H1 唯一性检查、TDK 字符长度打分评估)", foreground="gray").pack(anchor=tk.W, pady=(0, 15))

        # 按钮组
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=20)
        ttk.Button(btn_frame, text="🚀 选择映射表并开始执行", command=self._start_workflow).pack(side=tk.RIGHT, padx=5)
        
        self.root.mainloop()

    def get_tdk_and_validate(self, folder_path: Path, name: str) -> Tuple[str, str]:
        """提取 TDK 内容并根据长度打分建议"""
        try:
            word = name.split()[0].replace('.', '').lower()
            for file_path in folder_path.iterdir():
                if file_path.is_file():
                    fname = file_path.name.lower()
                    if fname.startswith(f"tdk-{word}") or fname == "tdk.docx":
                        doc = Document(file_path)
                        full_txt = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                        
                        advice = "未开启检查"
                        if self.var_seo_check.get():
                            title_len = 0
                            desc_len = 0
                            # 简单的正则提取打分 (容错处理)
                            t_match = re.search(r'(?i)title[:：]\s*(.*)', full_txt)
                            d_match = re.search(r'(?i)description[:：]\s*(.*)', full_txt)
                            
                            advices = []
                            if t_match:
                                title_len = len(t_match.group(1).strip())
                                if not (50 <= title_len <= 60): advices.append(f"Title({title_len}字符)不佳")
                            if d_match:
                                desc_len = len(d_match.group(1).strip())
                                if not (150 <= desc_len <= 160): advices.append(f"Desc({desc_len}字符)不佳")
                            
                            advice = "✅ TDK 长度优良" if not advices else "❌ " + ", ".join(advices)
                            if not t_match and not d_match: advice = "⚠️ 未识别到标准 Title/Desc 标签"
                            
                        return full_txt, advice
        except: return "读取失败", "-"
        return "缺失", "-"

    def _start_workflow(self):
        self.run_status = True
        self.root.destroy()

    def run(self):
        self._build_gui()
        if not self.run_status: return

        # 1. 加载映射表
        path = filedialog.askopenfilename(title="1. 选择【项目域名映射表】Excel", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try:
            df = pd.read_excel(path)
            self.domain_map = dict(zip(df.iloc[:,0].astype(str).str.strip(), df.iloc[:,1].astype(str).str.strip()))
        except Exception as e:
            messagebox.showerror("错误", f"读取 Excel 失败: {e}")
            return

        # 2. 选择目录
        root_dir_str = filedialog.askdirectory(title="2. 选择文章父文件夹")
        if not root_dir_str: return
        root_dir = Path(root_dir_str)
        
        # 3. 映射健康度校验
        folders = [f.name for f in root_dir.iterdir() if f.is_dir()]
        unmapped_test = [f for f in folders if f not in self.domain_map]
        if self.domain_map and len(unmapped_test) > len(folders) * 0.5:
            msg = f"⚠️ 严重警告：大半文件夹名称未在 Excel 第一列找到映射！\n例如：'{unmapped_test[0]}'\n这将导致外链清理失效。是否继续？"
            if not messagebox.askyesno("映射严重不匹配", msg): return
        
        out_f = filedialog.asksaveasfilename(title="4. 保存报告", defaultextension=".xlsx", initialfile="SEO审计修复报告_v6.5.xlsx")
        if not out_f: return

        config_dict = {
            'clean_links': self.var_clean_links.get(),
            'force_webp': self.var_force_webp.get(),
            'enable_seo_check': self.var_seo_check.get()
        }

        # 先审计
        self.execute_all(root_dir, config_dict, apply_fix=False)
        
        if messagebox.askyesno("确认修复", "初次扫描已完成！是否根据当前配置执行自动化修复？"):
            self.results.clear()
            self.execute_all(root_dir, config_dict, apply_fix=True)
            messagebox.showinfo("成功", "修复完成！请查看 Excel 报告。")
            
        pd.DataFrame([asdict(res) for res in self.results]).to_excel(out_f, index=False)

    def execute_all(self, root_dir: Path, config: dict, apply_fix: bool):
        for folder_path in root_dir.iterdir():
            if not folder_path.is_dir(): continue
            
            project_name = folder_path.name
            domain = self.domain_map.get(project_name)
            folder_files = [f.name.lower() for f in folder_path.iterdir() if f.is_file()]
            
            for file_path in folder_path.glob("*.docx"):
                if file_path.name.startswith(('~', 'TDK')): continue
                
                engine = SEOSuperEngineV65(file_path, domain, folder_files, config)
                h1_status, h_status, co_info, links, miss, logs = engine.process(apply_fix=apply_fix)
                tdk_content, tdk_advice = self.get_tdk_and_validate(folder_path, file_path.name)
                
                res = AuditResult(
                    project_folder=project_name,
                    file_name_link=f'=HYPERLINK("{file_path}", "{file_path.name}")',
                    domain_status=domain if domain else "❌ 未匹配映射(跳过外链清理)",
                    h1_status=h1_status,
                    heading_status=h_status,
                    missing_images=miss,
                    company_info=co_info,
                    tdk_content=tdk_content,
                    tdk_advice=tdk_advice,
                    all_links=links,
                    cleaned_links_count=str(engine.links_removed_count) if apply_fix else "等待修复",
                    logs=logs
                )
                self.results.append(res)

if __name__ == "__main__":
    SEOWorkflowManagerV65().run()