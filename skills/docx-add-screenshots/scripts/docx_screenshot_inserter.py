#!/usr/bin/env python3
"""
将浏览器截图批量插入 Word 文档。

通过 openCLI 获取浏览器截图，自动插入到 docx 文档的指定标题下方。
每张截图格式为：居中嵌入式图片 + 下方"图 N"标注。

用法:
    python3 docx_screenshot_inserter.py --docx <file.docx> --opencli <path> [options]

依赖: 仅 Python 标准库
"""

import argparse
import json
import os
import platform
import shutil
import struct
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.dom import minidom


# ======================== 平台检测 ========================

def detect_platform():
    """检测运行环境"""
    if os.path.exists("/proc/version"):
        with open("/proc/version", "r") as f:
            if "microsoft" in f.read().lower():
                return "wsl"
    if platform.system() == "Windows":
        return "windows"
    return "linux"


# ======================== openCLI 调用 ========================

class OpenCLIClient:
    """跨平台 openCLI 客户端"""

    def __init__(self, opencli, session="work", powershell=None):
        self.opencli = opencli
        self.session = session
        self.env = detect_platform()
        self.powershell = powershell
        self._bound_url = None

    def _run(self, args):
        """执行 openCLI 命令"""
        cmd_str = f"browser {self.session} {args}"

        if self.env == "wsl":
            ps = self.powershell or "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
            full = [ps, "-Command",
                    f"Set-Location C:\\; {self.opencli} {cmd_str}"]
        elif self.env == "windows":
            full = [self.opencli] + cmd_str.split()
        else:
            full = [self.opencli] + cmd_str.split()

        r = subprocess.run(full, capture_output=True, text=True, timeout=30)
        if r.returncode != 0 and r.stderr.strip():
            print(f"  [WARN] {r.stderr.strip()[:200]}")
        return r.stdout.strip()

    def bind(self):
        """绑定当前活跃 Chrome 标签页，返回 URL"""
        output = self._run("bind")
        try:
            info = json.loads(output)
            self._bound_url = info.get("url", "")
            return info
        except (json.JSONDecodeError, ValueError):
            self._bound_url = None
            return {"url": output, "raw": output}

    def get_bound_url(self):
        """获取当前绑定的 URL"""
        return self._bound_url

    def screenshot(self, output_path, width=1920, height=1080):
        """截图到指定路径（Windows 路径格式）"""
        self._run(f"screenshot {output_path} --width {width} --height {height}")


# ======================== PNG 工具 ========================

def get_png_dimensions(png_path):
    """从 PNG 文件头读取宽高（无需 PIL）"""
    with open(png_path, "rb") as f:
        f.read(8)   # PNG signature
        f.read(4)   # chunk length
        f.read(4)   # "IHDR"
        width = struct.unpack(">I", f.read(4))[0]
        height = struct.unpack(">I", f.read(4))[0]
    return width, height


# ======================== 文档处理 ========================

def get_paragraph_text(para):
    """提取段落纯文本"""
    parts = []
    for run in para.getElementsByTagName("w:r"):
        for t in run.getElementsByTagName("w:t"):
            if t.firstChild and t.firstChild.nodeType == t.firstChild.TEXT_NODE:
                parts.append(t.firstChild.data)
    return "".join(parts)


def get_paragraph_style(para):
    """获取段落样式 ID"""
    for ppr in para.getElementsByTagName("w:pPr"):
        for ps in ppr.getElementsByTagName("w:pStyle"):
            return ps.getAttribute("w:val")
    return None


def find_target_paragraphs(dom, heading_style=None, heading_text=None):
    """找到目标标题段落

    Args:
        heading_style: 按样式 ID 筛选（如 "17"）
        heading_text: 按文本匹配（逗号分隔的子串列表）

    Returns: [(para_node, text), ...]
    """
    results = []
    body = dom.getElementsByTagName("w:body")[0]

    # 解析文本过滤条件
    text_filters = None
    if heading_text:
        text_filters = [t.strip() for t in heading_text.split(",") if t.strip()]

    for para in body.getElementsByTagName("w:p"):
        # 样式过滤
        if heading_style:
            style = get_paragraph_style(para)
            if style != heading_style:
                continue

        text = get_paragraph_text(para)
        if not text.strip():
            continue

        # 文本过滤
        if text_filters:
            if not any(f in text for f in text_filters):
                continue

        results.append((para, text))

    return results


def get_next_rid(rels_dom):
    """获取下一个可用的 rId"""
    max_id = 0
    for rel in rels_dom.getElementsByTagName("Relationship"):
        rid = rel.getAttribute("Id")
        if rid.startswith("rId"):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass
    return f"rId{max_id + 1}"


def ensure_png_content_type(ct_dom):
    """确保 [Content_Types].xml 中声明了 png"""
    root = ct_dom.documentElement
    for elem in root.getElementsByTagName("Default"):
        if elem.getAttribute("Extension") == "png":
            return
    e = ct_dom.createElement("Default")
    e.setAttribute("Extension", "png")
    e.setAttribute("ContentType", "image/png")
    root.appendChild(e)


def insert_figure(para_node, image_filename, rid, width_emus, height_emus, fig_num):
    """在段落后插入图片 + 图片标注

    生成两个段落:
      1. 居中的 inline 图片
      2. 居中加粗的 "图 N" 标注
    """
    doc = para_node.ownerDocument
    doc_pr_id = fig_num + 200

    # --- 图片段落 ---
    img_p = doc.createElement("w:p")

    img_pPr = doc.createElement("w:pPr")
    img_jc = doc.createElement("w:jc")
    img_jc.setAttribute("w:val", "center")
    img_pPr.appendChild(img_jc)
    img_p.appendChild(img_pPr)

    img_r = doc.createElement("w:r")
    drawing = doc.createElement("w:drawing")
    inline = doc.createElement("wp:inline")
    for attr, val in [("distT", "0"), ("distB", "0"),
                       ("distL", "0"), ("distR", "0")]:
        inline.setAttribute(attr, val)

    extent = doc.createElement("wp:extent")
    extent.setAttribute("cx", str(width_emus))
    extent.setAttribute("cy", str(height_emus))
    inline.appendChild(extent)

    docPr = doc.createElement("wp:docPr")
    docPr.setAttribute("id", str(doc_pr_id))
    docPr.setAttribute("name", f"Figure {fig_num}")
    inline.appendChild(docPr)

    graphic = doc.createElement("a:graphic")
    graphicData = doc.createElement("a:graphicData")
    graphicData.setAttribute("uri",
        "http://schemas.openxmlformats.org/drawingml/2006/picture")

    pic = doc.createElement("pic:pic")
    nvPicPr = doc.createElement("pic:nvPicPr")
    cNvPr = doc.createElement("pic:cNvPr")
    cNvPr.setAttribute("id", str(doc_pr_id))
    cNvPr.setAttribute("name", image_filename)
    nvPicPr.appendChild(cNvPr)
    nvPicPr.appendChild(doc.createElement("pic:cNvPicPr"))
    pic.appendChild(nvPicPr)

    blipFill = doc.createElement("pic:blipFill")
    blip = doc.createElement("a:blip")
    blip.setAttribute("r:embed", rid)
    blipFill.appendChild(blip)
    stretch = doc.createElement("a:stretch")
    stretch.appendChild(doc.createElement("a:fillRect"))
    blipFill.appendChild(stretch)
    pic.appendChild(blipFill)

    spPr = doc.createElement("pic:spPr")
    xfrm = doc.createElement("a:xfrm")
    ext = doc.createElement("a:ext")
    ext.setAttribute("cx", str(width_emus))
    ext.setAttribute("cy", str(height_emus))
    xfrm.appendChild(ext)
    spPr.appendChild(xfrm)
    prstGeom = doc.createElement("a:prstGeom")
    prstGeom.setAttribute("prst", "rect")
    prstGeom.appendChild(doc.createElement("a:avLst"))
    spPr.appendChild(prstGeom)
    pic.appendChild(spPr)

    graphicData.appendChild(pic)
    graphic.appendChild(graphicData)
    inline.appendChild(graphic)
    drawing.appendChild(inline)
    img_r.appendChild(drawing)
    img_p.appendChild(img_r)

    # --- 标注段落 ---
    cap_p = doc.createElement("w:p")

    cap_pPr = doc.createElement("w:pPr")
    cap_jc = doc.createElement("w:jc")
    cap_jc.setAttribute("w:val", "center")
    cap_pPr.appendChild(cap_jc)
    cap_p.appendChild(cap_pPr)

    cap_r = doc.createElement("w:r")
    cap_rPr = doc.createElement("w:rPr")
    cap_b = doc.createElement("w:b")
    cap_rPr.appendChild(cap_b)
    cap_rPr.appendChild(doc.createElement("w:bCs"))
    cap_r.appendChild(cap_rPr)

    cap_t = doc.createElement("w:t")
    cap_t.setAttribute("xml:space", "preserve")
    cap_t.appendChild(doc.createTextNode(f"图 {fig_num}"))
    cap_r.appendChild(cap_t)
    cap_p.appendChild(cap_r)

    # 插入到目标段落之后：先插标注，再插图片（因为 insertAfter 是逆序）
    parent = para_node.parentNode
    ref = para_node.nextSibling
    if ref:
        parent.insertBefore(cap_p, ref)
        parent.insertBefore(img_p, cap_p)
    else:
        parent.appendChild(img_p)
        parent.appendChild(cap_p)


# ======================== 主流程 ========================

def main():
    parser = argparse.ArgumentParser(
        description="将浏览器截图插入 Word 文档")
    parser.add_argument("--docx", required=True, help="目标 docx 文件")
    parser.add_argument("--opencli", required=True, help="opencli 路径")
    parser.add_argument("--session", default="work", help="session 名")
    parser.add_argument("--powershell", default=None, help="powershell.exe 路径（WSL）")
    parser.add_argument("--width", type=int, default=1920, help="截图宽度")
    parser.add_argument("--height", type=int, default=1080, help="截图高度")
    parser.add_argument("--output-dir", default=None, help="截图保存目录")
    parser.add_argument("--heading-style", default=None, help="目标标题 style ID")
    parser.add_argument("--heading-text", default=None,
                        help="按文本匹配标题（逗号分隔子串）")
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"[ERROR] 文件不存在: {docx_path}")
        sys.exit(1)

    # 备份
    backup_path = docx_path.with_suffix(".backup.docx")
    if not backup_path.exists():
        shutil.copy2(docx_path, backup_path)
        print(f"已备份: {backup_path}")
    else:
        print(f"备份已存在: {backup_path}")

    # 截图目录
    screenshot_dir = Path(args.output_dir) if args.output_dir else docx_path.parent / "screenshots"
    os.makedirs(screenshot_dir, exist_ok=True)

    # 工作目录
    work_dir = docx_path.parent / f"._{docx_path.stem}_work"
    if work_dir.exists():
        shutil.rmtree(work_dir)

    # 解包
    print(f"解包文档: {docx_path}")
    with zipfile.ZipFile(docx_path, "r") as zf:
        zf.extractall(work_dir)

    # 解析 XML
    dom = minidom.parse(str(work_dir / "word" / "document.xml"))
    rels_dom = minidom.parse(str(work_dir / "word" / "_rels" / "document.xml.rels"))
    ct_dom = minidom.parse(str(work_dir / "[Content_Types].xml"))

    # 查找目标段落
    headings = find_target_paragraphs(dom, args.heading_style, args.heading_text)
    if not headings:
        print("[ERROR] 未找到匹配的标题。尝试指定 --heading-style 或 --heading-text")
        shutil.rmtree(work_dir)
        sys.exit(1)

    print(f"\n找到 {len(headings)} 个目标:")
    for _, text in headings:
        print(f"  - {text[:60]}")

    # 初始化 openCLI
    client = OpenCLIClient(args.opencli, args.session, args.powershell)

    print(f"\n{'='*60}")
    print(f"  半自动截图模式")
    print(f"  操作: 导航到界面 → 回车截图 | r=重绑 | s=跳过 | q=退出")
    print(f"{'='*60}\n")

    # 创建临时截图目录（Windows 可访问）
    tmp_screenshot_dir = Path("/tmp") / "docx_screenshots"
    if detect_platform() == "windows":
        tmp_screenshot_dir = Path(os.environ.get("TEMP", "C:\\Temp")) / "docx_screenshots"
    os.makedirs(tmp_screenshot_dir, exist_ok=True)

    fig_num = 0
    skipped = []

    for i, (para_node, text) in enumerate(headings, 1):
        print(f"[{i}/{len(headings)}] {text[:55]}")

        while True:
            # bind 到当前活跃标签页
            info = client.bind()
            bound_url = info.get("url", "unknown")
            print(f"  当前绑定: {bound_url}")

            user_input = input("  → 回车截图 | r=重新绑定 | s=跳过 | q=退出: ").strip().lower()

            if user_input == "q":
                print("\n用户中断，保存已完成的部分...")
                # 跳到保存步骤
                break
            elif user_input == "s":
                print("  已跳过\n")
                skipped.append(text[:40])
                break
            elif user_input == "r":
                print("  重新绑定...")
                continue
            else:
                # 回车 = 确认并截图
                break

        if user_input == "q":
            break
        if user_input == "s" or user_input == "r":
            continue

        # 截图
        tmp_filename = f"fig_{fig_num + 1}.png"
        win_screenshot = str(tmp_screenshot_dir / tmp_filename).replace("/", "\\")

        # 根据平台构造截图路径
        if detect_platform() == "wsl":
            # WSL: openCLI 输出 Windows 路径，需要转换
            win_path_for_opencli = f"C:\\Users\\Public\\{tmp_filename}"
            client.screenshot(win_path_for_opencli, args.width, args.height)
            wsl_path = f"/mnt/c/Users/Public/{tmp_filename}"
        else:
            client.screenshot(win_screenshot, args.width, args.height)
            wsl_path = str(tmp_screenshot_dir / tmp_filename)

        if not os.path.exists(wsl_path):
            print(f"  [ERROR] 截图未生成: {wsl_path}")
            skipped.append(text[:40])
            continue

        # 获取尺寸
        img_w, img_h = get_png_dimensions(wsl_path)
        width_emus = int(6.5 * 914400)
        height_emus = int(width_emus * img_h / img_w)

        # 复制到 media 目录
        media_dir = work_dir / "word" / "media"
        os.makedirs(media_dir, exist_ok=True)
        fig_num += 1
        image_filename = f"figure_{fig_num}.png"
        shutil.copy2(wsl_path, media_dir / image_filename)

        # 保存到输出目录
        shutil.copy2(wsl_path, screenshot_dir / image_filename)

        # 注册 relationship
        rid = get_next_rid(rels_dom)
        rel = rels_dom.createElement("Relationship")
        rel.setAttribute("Id", rid)
        rel.setAttribute("Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel.setAttribute("Target", f"media/{image_filename}")
        rels_dom.documentElement.appendChild(rel)

        ensure_png_content_type(ct_dom)

        # 插入图片 + 标注
        insert_figure(para_node, image_filename, rid, width_emus, height_emus, fig_num)

        print(f"  已插入: 图 {fig_num} ({img_w}x{img_h})\n")

    # 保存 XML
    for path, dom_obj in [
        (work_dir / "word" / "document.xml", dom),
        (work_dir / "word" / "_rels" / "document.xml.rels", rels_dom),
        (work_dir / "[Content_Types].xml", ct_dom),
    ]:
        with open(str(path), "w", encoding="utf-8") as f:
            f.write(dom_obj.toxml())

    # 打包
    print(f"重新打包文档...")
    with zipfile.ZipFile(str(docx_path), "w", zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(work_dir):
            for file in files:
                fp = os.path.join(root_dir, file)
                zf.write(fp, os.path.relpath(fp, work_dir))

    shutil.rmtree(work_dir)

    print(f"\n完成！")
    print(f"  文档: {docx_path}")
    print(f"  已插入: {fig_num} 张截图（图 1 ~ 图 {fig_num}）")
    print(f"  截图保存: {screenshot_dir}")
    if skipped:
        print(f"  已跳过: {', '.join(skipped)}")


if __name__ == "__main__":
    main()
