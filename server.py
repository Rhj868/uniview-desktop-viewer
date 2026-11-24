# -*- coding: utf-8 -*-
"""
UniView - 本地文件统一查看器 (桌面窗口版)
------------------------------------------------
形态：双击 start.bat 启动 -> 本机只监听 127.0.0.1 的小型服务
      -> 用 Edge/Chrome --app 模式打开一个无地址栏的独立窗口。

浏览器网页做不到、而这里能做到的：
  1. 按路径读取任意本地文件（不需要用户每次手选）
  2. 扫描目录树、跨格式全文搜索（含 docx/xlsx/pptx 正文抽取）
  3. 把修改直接写回原文件（保持原编码，首次保存自动备份）
  4. 调用/嵌入本机 WPS，让 Office 文档显示在自己的窗口里

零第三方安装：只用 Python 标准库 + 本机已有的 Edge/Chrome + 已装 WPS。
"""

import argparse
import base64
import ctypes
import glob
import html as html_lib
import json
import mimetypes
import os
import posixpath
import re
import secrets
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import urllib.request
import tempfile
import webbrowser
from ctypes import wintypes
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

if getattr(sys, "frozen", False):
    # PyInstaller 单文件：__file__ 在只读的解包临时目录里。
    # 资源（index.html/vendor/图标）从那里读，可写数据统一放 %APPDATA%\UniView。
    ROOT = sys._MEIPASS
else:
    ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.environ.get("APPDATA") or ROOT, "UniView")
try:
    os.makedirs(DATA_DIR, exist_ok=True)
except OSError:
    DATA_DIR = ROOT
TOKEN = secrets.token_urlsafe(16)
TITLE_TAG = "UniView-" + secrets.token_hex(4)   # 仅兼容旧窗口；新窗口标题固定为干净的 "UniView"
HTML_KEY = secrets.token_urlsafe(12)            # HTML 沙箱视图专用钥匙，与 API 令牌分开
PORT = 0
LOG_PATH = os.path.join(DATA_DIR, "uniview.log")
RUNTIME_PATH = os.path.join(DATA_DIR, ".runtime.json")   # 单实例：记录当前服务的端口与令牌
SESSION_PATH = os.path.join(DATA_DIR, "session.json")    # 会话：上次打开的标签，下次启动自动恢复
ROOTS_PATH = os.path.join(DATA_DIR, "roots.json")        # 「浏览其他文件夹…」选过的目录（后端持久化）
MAX_CUSTOM_ROOTS = 15                                   # 下拉最多保留最近 15 个记住的目录，超出挤掉最旧的
BROWSE_ACTIVE = threading.Event()                       # 原生文件夹对话框弹出期间，暂停主窗口抢焦点

# 「打开方式」后续文件不再另开窗口，而是排队等当前窗口来取，前端取到后开成新标签
PENDING = []
PENDING_LOCK = threading.Lock()
STATE = {"seen": 0.0}   # 窗口最近一次轮询时间；超过阈值说明窗口已关，需要重新开一个


def log(msg):
    """把关键步骤写到 uniview.log，便于真机出问题时回看，不依赖控制台。"""
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(time.strftime("%H:%M:%S") + " " + str(msg) + "\n")
    except Exception:
        pass

# ---------------------------------------------------------------- 格式分类

TEXT_EXT = {
    ".txt", ".md", ".markdown", ".csv", ".tsv", ".log", ".ini", ".json",
    ".xml", ".yaml", ".yml", ".html", ".htm", ".css", ".js", ".ts", ".jsx",
    ".tsx", ".py", ".java", ".c", ".cpp", ".h", ".cs", ".go", ".rs", ".rb",
    ".php", ".sql", ".sh", ".bat", ".ps1", ".toml", ".vue", ".svg",
}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tif", ".tiff"}
OFFICE_EXT = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".wps", ".et", ".dps"}
PDF_EXT = {".pdf"}
RTF_EXT = {".rtf"}
VIEWABLE = TEXT_EXT | IMAGE_EXT | OFFICE_EXT | PDF_EXT | RTF_EXT

MAX_SEARCH_FILES = 400
MAX_SEARCH_BYTES = 3 * 1024 * 1024


THIS_PC = "__UNIVIEW_THIS_PC__"   # 「此电脑」虚拟根：/api/tree 见到它就返回磁盘列表


def read_custom_roots():
    """读记住的目录，顺手剔除已被删除/移动的目录。
    必须存在后端：服务每次启动端口随机，浏览器 localStorage 按端口隔离，跨启动读不到。"""
    try:
        with open(ROOTS_PATH, encoding="utf-8") as f:
            d = json.load(f)
        arr = d.get("roots", []) if isinstance(d, dict) else []
        return [p for p in arr if isinstance(p, str) and os.path.isdir(p)][-MAX_CUSTOM_ROOTS:]
    except (OSError, ValueError):
        return []


def write_custom_roots(arr):
    try:
        with open(ROOTS_PATH, "w", encoding="utf-8") as f:
            json.dump({"roots": arr[-MAX_CUSTOM_ROOTS:]}, f, ensure_ascii=False)
    except OSError:
        pass


def list_drives():
    """「此电脑」下的磁盘，格式与 /api/tree 的 items 完全一致。
    这样前端不用改树渲染：盘符就是一个普通的可展开目录，点开后逐层下钻。"""
    items = []
    if os.name != "nt":
        return items
    try:
        k32 = ctypes.windll.kernel32
        buf = ctypes.create_unicode_buffer(512)
        n = k32.GetLogicalDriveStringsW(511, buf)
        raw = ctypes.wstring_at(ctypes.addressof(buf), n) if n else ""
        for d in raw.split("\0"):
            d = d.strip()
            if not d:
                continue
            # 0 未知 / 1 无根目录 / 5 光驱（多为空）跳过，避免点开就报错
            if k32.GetDriveTypeW(d) not in (2, 3, 4, 6):
                continue
            label = ""
            try:
                vbuf = ctypes.create_unicode_buffer(256)
                if k32.GetVolumeInformationW(d, vbuf, 256, None, None, None, None, 0):
                    label = vbuf.value
            except Exception:
                pass
            items.append({"name": (label + " (" + d[:2] + ")") if label else d,
                          "path": d, "dir": True, "viewable": True, "ext": "", "size": 0})
    except Exception as ex:
        log("DRIVES error " + repr(ex))
    return items


def browse_folder(title="选择要在 UniView 中浏览的文件夹"):
    """Windows 原生「选择文件夹」对话框：从「此电脑」展开 C:/D: 盘逐层挑目录。
    返回选中的绝对路径；取消或失败返回空串。纯 ctypes，无需第三方库。"""
    if os.name != "nt":
        return ""

    class BROWSEINFOW(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),
            ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", ctypes.c_void_p),
            ("lpszTitle", wintypes.LPWSTR),
            ("ulFlags", wintypes.UINT),
            ("lpfn", ctypes.c_void_p),
            ("lParam", ctypes.c_void_p),
            ("iImage", ctypes.c_int),
        ]

    ole32, shell32 = ctypes.windll.ole32, ctypes.windll.shell32
    user32 = ctypes.windll.user32
    ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
    ole32.CoInitializeEx(None, 0x2)     # COINIT_APARTMENTTHREADED
    # 对话框由「无窗口的服务进程」弹出，默认拿不到前台、也没有父窗口，
    # 会被 UniView 主窗口盖住。用 BFFM_INITIALIZED 回调把它设为置顶并抢到前台。
    BFFCALLBACK = ctypes.WINFUNCTYPE(ctypes.c_int, wintypes.HWND, wintypes.UINT,
                                     wintypes.LPARAM, wintypes.LPARAM)
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND] + [ctypes.c_int] * 4 + [wintypes.UINT]
    user32.SetWindowPos.restype = wintypes.BOOL
    user32.BringWindowToTop.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.keybd_event.argtypes = [ctypes.c_ubyte, ctypes.c_ubyte,
                                   wintypes.DWORD, ctypes.c_void_p]
    user32.IsWindow.argtypes = [wintypes.HWND]
    user32.IsWindow.restype = wintypes.BOOL

    def steal_foreground(hwnd):
        """对话框刚初始化时抢前台常被系统拒绝；等它进入模态循环后再试一次，
        用一次 ALT 模拟解除 Windows 前台锁定，让弹窗出现后就能直接键盘操作。"""
        time.sleep(0.3)
        try:
            if not user32.IsWindow(hwnd):
                return
            user32.keybd_event(0x12, 0, 0, None)
            user32.keybd_event(0x12, 0, 2, None)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        except Exception as ex:
            log("BROWSE foreground error " + repr(ex))

    @BFFCALLBACK
    def on_dialog_event(hwnd, msg, _lp, _data):
        if msg == 1:                     # BFFM_INITIALIZED
            try:
                # HWND_TOPMOST=-1；SWP_NOSIZE|SWP_NOMOVE|SWP_SHOWWINDOW
                user32.SetWindowPos(hwnd, ctypes.c_void_p(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040)
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
                threading.Thread(target=steal_foreground, args=(hwnd,), daemon=True).start()
            except Exception as ex:
                log("BROWSE topmost error " + repr(ex))
        return 0

    BROWSE_ACTIVE.set()
    try:
        shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFOW)]
        shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
        shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]
        shell32.SHGetPathFromIDListW.restype = wintypes.BOOL
        disp = ctypes.create_unicode_buffer(260)
        bi = BROWSEINFOW()
        # 绑定 UniView 窗口为 owner：模态对话框不会被 owner 盖住，且自动禁用主窗口
        try:
            owner = find_uniview_window() if callable(globals().get("find_uniview_window")) else None
            bi.hwndOwner = ctypes.c_void_p(owner) if owner else None
        except Exception:
            bi.hwndOwner = None
        bi.pidlRoot = None
        bi.pszDisplayName = ctypes.addressof(disp)   # 有效缓冲区，Shell 会写入显示名
        bi.lpszTitle = title
        # RETURNONLYFSDIRS | EDITBOX(可手填路径) | NEWDIALOGSTYLE(可缩放/新建文件夹)
        bi.ulFlags = 0x0001 | 0x0010 | 0x0040
        bi.lpfn = ctypes.cast(on_dialog_event, ctypes.c_void_p)
        bi.lParam = None
        bi.iImage = 0
        pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
        if not pidl:
            return ""
        try:
            out = ctypes.create_unicode_buffer(1024)
            return out.value if shell32.SHGetPathFromIDListW(pidl, out) else ""
        finally:
            ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
            ole32.CoTaskMemFree(pidl)
    except Exception as ex:
        log("BROWSE error " + repr(ex))
        return ""
    finally:
        BROWSE_ACTIVE.clear()
        ole32.CoUninitialize()


def sniff_image_type(data):
    """md 里引用的图片可能没有扩展名，按文件头补一个正确的 Content-Type，
    否则浏览器拿到 application/octet-stream 不会把 <img> 画出来。"""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "image/png"
    if data[:3] == b"\xff\xd8\xff":
        return "image/jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    if data[:2] == b"BM":
        return "image/bmp"
    return None


def kind_of(path):
    e = os.path.splitext(path)[1].lower()
    if e in {".html", ".htm"}:
        return "html"
    if e == ".rtf":
        return "rtf"
    if e in TEXT_EXT:
        return "text"
    if e in IMAGE_EXT:
        return "image"
    if e in PDF_EXT:
        return "pdf"
    if e in OFFICE_EXT:
        return "office"
    return "other"


def _b64e(s):
    return base64.urlsafe_b64encode(s.encode("utf-8")).decode("ascii").rstrip("=")


def _b64d(s):
    s += "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s.encode("ascii")).decode("utf-8")


def html_view_url(path):
    """隔离渲染地址：钥匙放在路径里，iframe 沙箱因此拿不到真正的 API 令牌。"""
    return "/html/%s/%s" % (HTML_KEY, _b64e(os.path.abspath(path)))


def build_html_document(path, base_url):
    """把本地 HTML 变成「只认本机、绝不出网」的文档。

    <base> 让相对资源走同目录代理；强制 utf-8；去掉 meta refresh；
    拦下指向外站的链接点击。真正的联网封锁由响应头 CSP 完成。
    """
    with open(path, "rb") as f:
        data = f.read()
    text, _ = decode_bytes(data)
    guard = (
        '<meta charset="utf-8">'
        '<base href="' + base_url + '/">'
        '<script>document.addEventListener("click",function(e){'
        'var a=e.target&&e.target.closest&&e.target.closest("a");'
        'if(a&&a.href&&!/^https?:\\/\\/127\\.0\\.0\\.1(:[0-9]+)?\\//.test(a.href))'
        '{e.preventDefault();}},true);</script>'
    )
    text = re.sub(r"<meta[^>]+http-equiv\s*=\s*['\"]?refresh[^>]*>", "", text, flags=re.I)
    m = re.search(r"<head[^>]*>", text, re.I)
    if m:
        text = text[:m.end()] + guard + text[m.end():]
    else:
        m = re.search(r"<html[^>]*>", text, re.I)
        if m:
            text = text[:m.end()] + "<head>" + guard + "</head>" + text[m.end():]
        else:
            text = guard + text
    return text.encode("utf-8")


def decode_bytes(data):
    """编码识别：BOM -> UTF-8 -> GB18030(覆盖 GBK/GB2312) -> cp1252 -> latin-1"""
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8-sig"
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16"), "utf-16"
    for enc in ("utf-8", "gb18030", "big5", "cp1252"):
        try:
            return data.decode(enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("latin-1"), "latin-1"


def read_text(path, limit=MAX_SEARCH_BYTES):
    with open(path, "rb") as f:
        data = f.read() if limit is None else f.read(limit)
    return decode_bytes(data)


def is_probably_binary(data):
    if b"\x00" in data[:4096]:
        return True
    return False


# ---------------------------------------------------------------- 全文抽取

def _xml_texts(data, tags):
    """从 OOXML 片段里取指定标签的文本（纯标准库，不依赖 python-docx 等包）。"""
    import xml.etree.ElementTree as ET
    root = ET.fromstring(data)
    return ["".join(el.itertext()) for el in root.iter()
            if el.tag.rsplit("}", 1)[-1] in tags]


def extract_text(path):
    """跨格式全文搜索用：抽取纯文本。docx/xlsx/pptx 本质是 zip+XML，直接用标准库读。"""
    e = os.path.splitext(path)[1].lower()
    try:
        if e == ".rtf":
            with open(path, "rb") as f:
                return rtf_plain_text(f.read())
        if e in TEXT_EXT:
            return read_text(path)[0]
        if e in (".docx", ".xlsx", ".pptx"):
            import zipfile
            out = []
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
                if e == ".docx":
                    for n in names:
                        if n == "word/document.xml" or re.match(r"word/(header|footer)\d*\.xml$", n):
                            out += _xml_texts(z.read(n), {"t"})
                elif e == ".pptx":
                    for n in sorted(n for n in names
                                    if re.match(r"ppt/slides/slide\d+\.xml$", n)):
                        out += _xml_texts(z.read(n), {"t"})
                else:
                    for n in sorted(n for n in names
                                    if re.match(r"xl/(sharedStrings\.xml|worksheets/sheet\d+\.xml)$", n)):
                        out += _xml_texts(z.read(n), {"t", "v"})
            return "\n".join(x for x in out if x and x.strip())
    except Exception:
        return ""
    return ""


# ---------------------------------------------------------------- PPTX 内置预览
# 关键点：不把整个 pptx 传给浏览器（实测用户文件 297MB，其中 270MB 是嵌入对象），
# 只在服务端读 slide XML / rels，图片由 /api/office/media 按需单独流式取。

# 只列浏览器能直接解码的格式；TIFF/EMF/WMF 放进来只会显示成裂图，
# 交给下面的分支给出「请用 WPS 嵌入查看」的明确提示。
PPT_IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg"}
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _ln(tag):
    """取 XML 标签的本地名，忽略命名空间前缀。"""
    return tag.rsplit("}", 1)[-1]


def _ppt_rels(z, slide_name):
    base = posixpath.dirname(slide_name)
    rel = base + "/_rels/" + posixpath.basename(slide_name) + ".rels"
    out = {}
    try:
        import xml.etree.ElementTree as ET
        for el in ET.fromstring(z.read(rel)):
            rid, tgt = el.get("Id"), el.get("Target") or ""
            if tgt.startswith("/"):
                norm = tgt.lstrip("/")
            else:
                norm = posixpath.normpath(posixpath.join(base, tgt))
            out[rid] = norm
    except Exception:
        pass
    return out


def _ppt_run(r, slide_cx):
    props, txt = None, []
    for c in r:
        n = _ln(c.tag)
        if n == "rPr":
            props = c
        elif n == "t":
            txt.append(html_lib.escape(c.text or ""))
        elif n == "br":
            txt.append("<br>")
        elif n == "tab":
            txt.append("&nbsp;&nbsp;&nbsp;&nbsp;")
    css = []
    if props is not None:
        if props.get("b") in ("1", "true"):
            css.append("font-weight:700")
        if props.get("i") in ("1", "true"):
            css.append("font-style:italic")
        if props.get("u") not in (None, "none"):
            css.append("text-decoration:underline")
        sz = props.get("sz") or ""
        if sz.isdigit() and slide_cx:
            # sz 是「百分点」：pt = sz/100，px = pt*96/72；再换算为「占画布宽度的百分比」
            # cqw = px / nativeW * 100，nativeW = slide_cx/9525 → cqw = sz*12700/slide_cx
            css.append("font-size:%.3fcqw" % (int(sz) * 12700.0 / slide_cx))
        for c in props.iter():
            if _ln(c.tag) == "srgbClr" and c.get("val"):
                css.append("color:#" + c.get("val"))
                break
    return ";".join(css), "".join(txt)


def _ppt_para_html(p, slide_cx):
    pPr = next((c for c in p if _ln(c.tag) == "pPr"), None)
    align, bullet = "", ""
    if pPr is not None:
        align = {"ctr": "center", "r": "right", "just": "justify"}.get(pPr.get("algn") or "", "")
        names = {_ln(c.tag) for c in pPr}
        if not {"buNone"} & names and names & {"buChar", "buAutoNum", "buBlip"}:
            bullet = "&bull;&nbsp;"
    inner = ""
    for n in p:
        k = _ln(n.tag)
        if k == "r":
            css, t = _ppt_run(n, slide_cx)
            inner += ('<span style="%s">%s</span>' % (css, t)) if (css and t) else t
        elif k == "fld":
            inner += "".join(html_lib.escape(c.text or "") for c in n.iter() if _ln(c.tag) == "t")
        elif k == "br":
            inner += "<br>"
    if not inner.strip():
        return ""
    style = ' style="text-align:%s"' % align if align else ""
    return "<p%s>%s%s</p>" % (style, bullet, inner)


def _ppt_pic_tag(pic, rels, ppt_path):
    """只产出 <img>/占位；位置与大小由外层绝对定位容器负责。"""
    rid = ""
    for b in pic.iter():
        if _ln(b.tag) != "blip":
            continue
        for k, v in b.attrib.items():
            if k.rsplit("}", 1)[-1] == "embed":
                rid = v
                break
        if rid:
            break
    tgt = rels.get(rid, "")
    if not tgt:
        return '<div class="ppt-noimg">［图片关联丢失］</div>'
    ext = os.path.splitext(tgt)[1].lower()
    if ext not in PPT_IMG_EXT:
        return ('<div class="ppt-noimg">［%s 图片浏览器不支持，请用「用 WPS 嵌入」查看］</div>'
                % html_lib.escape(ext or "未知"))
    src = "/api/office/media?" + urllib.parse.urlencode(
        {"path": ppt_path, "name": tgt, "t": TOKEN})
    return ('<img class="ppt-photo" loading="lazy" src="%s" alt="幻灯片图片">'
            % html_lib.escape(src, quote=True))


def _ppt_table_html(gf, slide_cx):
    tbl = next((c for c in gf.iter() if _ln(c.tag) == "tbl"), None)
    if tbl is None:
        return ""
    rows = []
    for tr in tbl:
        if _ln(tr.tag) != "tr":
            continue
        cells = []
        for tc in tr:
            if _ln(tc.tag) != "tc":
                continue
            body = "".join(_ppt_para_html(p, slide_cx) for p in tc.iter()
                           if _ln(p.tag) == "p")
            cells.append("<td>%s</td>" % (body or "&nbsp;"))
        if cells:
            rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<table class="csv">' + "".join(rows) + "</table>" if rows else ""


def _ppt_content(el, rels, ppt_path, slide_cx):
    n = _ln(el.tag)
    if n == "sp":
        tx = next((c for c in el if _ln(c.tag) == "txBody"), None)
        if tx is None:
            return ""
        body = "".join(x for x in (_ppt_para_html(p, slide_cx) for p in tx
                                   if _ln(p.tag) == "p") if x)
        if not body.strip():
            return ""
        bodyPr = next((c for c in tx if _ln(c.tag) == "bodyPr"), None)
        anchor = bodyPr.get("anchor") if bodyPr is not None else ""
        if anchor == "ctr":
            return '<div class="ppt-vc">%s</div>' % body
        if anchor == "b":
            return '<div class="ppt-vb">%s</div>' % body
        return body
    if n == "pic":
        return _ppt_pic_tag(el, rels, ppt_path)
    if n == "graphicFrame":
        return _ppt_table_html(el, slide_cx)
    return ""


def _ppt_xfrm_read(x):
    d = {"ox": 0, "oy": 0, "cx": 0, "cy": 0, "chx": 0, "chy": 0,
         "chcx": 0, "chcy": 0, "rot": 0}
    for e in x:
        k = _ln(e.tag)
        if k == "off":
            d["ox"] = int(e.get("x") or 0)
            d["oy"] = int(e.get("y") or 0)
        elif k == "ext":
            d["cx"] = int(e.get("cx") or 0)
            d["cy"] = int(e.get("cy") or 0)
        elif k == "chOff":
            d["chx"] = int(e.get("x") or 0)
            d["chy"] = int(e.get("y") or 0)
        elif k == "chExt":
            d["chcx"] = int(e.get("cx") or 0)
            d["chcy"] = int(e.get("cy") or 0)
    if x.get("rot"):
        d["rot"] = int(x.get("rot"))
    return d


def _ppt_xfrm(el):
    for c in el:
        k = _ln(c.tag)
        if k in ("spPr", "grpSpPr"):
            for d in c:
                if _ln(d.tag) == "xfrm":
                    return _ppt_xfrm_read(d)
        elif k == "xfrm":
            return _ppt_xfrm_read(c)
    return None


def _ppt_walk(el, tr, sw, sh, rels, ppt_path, out):
    r"""把一个形状按真实坐标换算成百分比绝对定位；组合按 chOff/chExt 递归映射。

    tr = (sx, sy, tx, ty)：局部 EMU -> 幻灯片 EMU 的缩放+平移。
    """
    n = _ln(el.tag)
    sx, sy, tx, ty = tr
    if n == "grpSp":
        g = _ppt_xfrm(el)
        ctr = tr
        if g and g["chcx"] and g["chcy"]:
            nsx = sx * g["cx"] / g["chcx"]
            nsy = sy * g["cy"] / g["chcy"]
            ntx = tx + sx * g["ox"] - nsx * g["chx"]
            nty = ty + sy * g["oy"] - nsy * g["chy"]
            ctr = (nsx, nsy, ntx, nty)
        for ch in el:
            if _ln(ch.tag) in ("sp", "pic", "graphicFrame", "grpSp"):
                _ppt_walk(ch, ctr, sw, sh, rels, ppt_path, out)
        return
    if n not in ("sp", "pic", "graphicFrame"):
        return
    content = _ppt_content(el, rels, ppt_path, sw)
    if not content:
        return
    xf = _ppt_xfrm(el)
    if not xf or not xf["cx"] or not xf["cy"]:
        out.append("loose:" + content)   # 没有坐标就退回普通流，别把内容丢了
        return
    left = (tx + sx * xf["ox"]) * 100.0 / sw
    top = (ty + sy * xf["oy"]) * 100.0 / sh
    w = sx * xf["cx"] * 100.0 / sw
    h = sy * xf["cy"] * 100.0 / sh
    style = "left:%.3f%%;top:%.3f%%;width:%.3f%%;height:%.3f%%" % (left, top, w, h)
    rot = xf.get("rot", 0) / 60000.0
    if rot:
        style += ";transform:rotate(%.2fdeg)" % rot
    out.append("shape:" + '<div class="ppt-shape" style="%s">%s</div>' % (style, content))


def build_pptx_preview(path):
    import zipfile
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        sw, sh = 0, 0
        if "ppt/presentation.xml" in names:
            try:
                for el in ET.fromstring(z.read("ppt/presentation.xml")).iter():
                    if _ln(el.tag) == "sldSz":
                        sw = int(el.get("cx") or 0)
                        sh = int(el.get("cy") or 0)
                        break
            except Exception:
                pass
        if not sw or not sh:
            sw, sh = 12192000, 6858000          # 16:9 默认
        slides = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                        key=lambda x: int(re.search(r"slide(\d+)", x).group(1)))
        blocks, images, skipped = [], 0, 0
        for idx, sn in enumerate(slides, 1):
            rels = _ppt_rels(z, sn)
            try:
                root = ET.fromstring(z.read(sn))
            except Exception as ex:
                blocks.append('<div class="ppt-slide-wrap"><div class="no">第 %d 页</div>'
                              '<p>本页解析失败：%s</p></div>'
                              % (idx, html_lib.escape(str(ex))))
                continue
            tree = next((c for c in root.iter() if _ln(c.tag) == "spTree"), None)
            out = []
            if tree is not None:
                for ch in tree:
                    if _ln(ch.tag) in ("sp", "pic", "graphicFrame", "grpSp"):
                        _ppt_walk(ch, (1.0, 1.0, 0.0, 0.0), sw, sh, rels, path, out)
            shapes = "".join(x[6:] for x in out if x.startswith("shape:"))
            loose = "".join(x[6:] for x in out if x.startswith("loose:"))
            if shapes:
                canvas = ('<div class="ppt-slide" style="aspect-ratio:%d/%d">%s</div>'
                          % (sw, sh, shapes))
            else:
                canvas = ('<div class="ppt-slide ppt-blank" style="aspect-ratio:%d/%d">'
                          '<p style="color:var(--muted)">（本页无可显示内容）</p></div>'
                          % (sw, sh))
            loose_html = ('<div class="ppt-loose">%s</div>' % loose) if loose else ""
            images += canvas.count('class="ppt-photo"') + loose.count('class="ppt-photo"')
            skipped += canvas.count('class="ppt-noimg"') + loose.count('class="ppt-noimg"')
            blocks.append('<div class="ppt-slide-wrap"><div class="no">第 %d 页</div>%s%s</div>'
                          % (idx, canvas, loose_html))
    if not blocks:
        return {"html": '<div class="note-card"><h2>未找到幻灯片</h2></div>',
                "images": 0, "skipped": 0}
    return {"html": '<div class="ppt-deck">' + "".join(blocks) + "</div>",
            "images": images, "skipped": skipped}


# ---------------------------------------------------------------- RTF 解析（纯标准库）

RTF_CODECS = {1252: "cp1252", 1250: "cp1250", 1251: "cp1251", 1253: "cp1253",
              1254: "cp1254", 1255: "cp1255", 1256: "cp1256", 874: "cp874",
              932: "cp932", 936: "gbk", 949: "cp949", 950: "big5",
              65001: "utf-8", 437: "cp437", 850: "cp850"}
RTF_WINGDINGS = {0xF06C: "●", 0xF0B7: "▪", 0xF0A7: "▪", 0xF0D8: "➢",
                 0xF0FC: "✓", 0xF0FD: "✗", 0xF0A1: "❖"}
RTF_SKIP_DESTS = {
    "fonttbl", "stylesheet", "info", "title", "subject", "author", "manager",
    "company", "operator", "creatim", "revtim", "printim", "buptim", "comment",
    "doccomm", "hlinkbase", "xmlnstbl", "listtable", "listoverridetable",
    "rsidtbl", "generator", "filetbl", "revtbl", "userprops", "protusertbl",
    "latentstyles", "datastore", "themedata", "colorschememapping",
    "wgrffmtfilter", "pnseclvl", "pntxta", "pntxtb", "listtext", "listname",
    "nonshppict", "object", "objdata", "objalias", "objsect", "objclass",
    "objname", "objtime", "header", "footer", "headerl", "headerr", "footerl",
    "footerr", "footnote", "annotation", "atnid", "atnauthor", "fldinst",
    "datafield", "result", "picprop", "blipuid", "mmathPr", "mmath",
    "background", "xe", "tc", "tcn", "private", "password", "doctag", "panose",
    "falt", "fname", "shppict", "themedata", "xmlnsname", "shpinst", "shprslt",
}
# shppict 不能跳过：它里面才是我们要的 \pict；nonshppict 是给 Word 的重复副本
RTF_SKIP_DESTS.discard("shppict")


def _rtf_group_end(s, i):
    r"""i 指向 '{'，返回匹配 '}' 之后的位置；跳过 \{ \} \'hh 转义。"""
    depth, n = 0, len(s)
    while i < n:
        c = s[i]
        if c == "\\":
            i += 4 if (i + 1 < n and s[i + 1] == "'") else 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return n


def _rtf_group_word(s, i):
    r"""读 '{' 之后（允许 \* 前缀）的第一个控制字。"""
    j = i + 1
    if s.startswith("\\*", j):
        j += 2
        if j < len(s) and s[j] == " ":
            j += 1
    if j >= len(s) or s[j] != "\\":
        return ""
    m = re.match(r"[A-Za-z]+", s[j + 1:])
    return m.group(0) if m else ""


def _rtf_unit_end(s, i):
    r"""跳过一个「字符单元」：普通字符 / 控制字 / \'hh。用于 \uN 的 fallback。"""
    if i >= len(s):
        return i
    if s[i] != "\\":
        return i + 1
    if i + 1 < len(s) and s[i + 1] == "'":
        return i + 4
    k = i + 1
    while k < len(s) and s[k].isalpha():
        k += 1
    while k < len(s) and (s[k].isdigit() or s[k] == "-"):
        k += 1
    return max(k, i + 2)


def _rtf_pict_html(seg):
    if "\\bin" in seg:
        return '<div class="rtf-noimg">［二进制图片暂不支持显示］</div>'
    kw = "\\pngblip" if "\\pngblip" in seg else ("\\jpegblip" if "\\jpegblip" in seg else "")
    if not kw:
        return '<div class="rtf-noimg">［WMF/EMF 矢量图片暂不支持显示，请用 WPS 打开］</div>'
    fmt = "image/png" if kw == "\\pngblip" else "image/jpeg"
    body = seg[seg.find(kw) + len(kw):]
    body = re.sub(r"\{[^{}]*\}", "", body)          # 去掉 blipuid 等嵌套组
    body = re.sub(r"\\[A-Za-z]+-?\d*\s?", "", body)  # 去掉控制字
    hexs = "".join(re.findall(r"[0-9a-fA-F]", body))
    if len(hexs) % 2:
        hexs = hexs[:-1]
    try:
        blob = bytes.fromhex(hexs)
    except ValueError:
        return '<div class="rtf-noimg">［图片数据损坏］</div>'
    magic = b"\x89PNG\r\n\x1a\n" if fmt == "image/png" else b"\xff\xd8"
    k = blob.find(magic)
    if k > 0:
        blob = blob[k:]
    if fmt == "image/png":
        e = blob.find(b"IEND")
        if e > 0:
            blob = blob[:e + 8]
    else:
        e = blob.rfind(b"\xff\xd9")
        if e > 0:
            blob = blob[:e + 2]
    if len(blob) < 100 or len(blob) > 12 * 1024 * 1024:
        return '<div class="rtf-noimg">［图片过大或为空，请用 WPS 打开］</div>'
    w = re.search(r"\\picwgoal(\d+)", seg)
    style = ""
    if w:
        try:
            style = ' style="width:%dpx"' % min(760, max(24, int(w.group(1)) // 15))
        except ValueError:
            pass
    return ('<img class="rtf-img"%s src="data:%s;base64,%s" alt="文档图片">'
            % (style, fmt, base64.b64encode(blob).decode("ascii")))


def _rtf_char_css(st):
    css = []
    if st["bold"]:
        css.append("font-weight:700")
    if st["italic"]:
        css.append("font-style:italic")
    dec = []
    if st["underline"]:
        dec.append("underline")
    if st["strike"]:
        dec.append("line-through")
    if dec:
        css.append("text-decoration:" + " ".join(dec))
    if st["size"]:
        css.append("font-size:%gpt" % st["size"])
    if st["color"]:
        css.append("color:" + st["color"])
    if st["highlight"]:
        css.append("background-color:" + st["highlight"])
    if st["caps"]:
        css.append("text-transform:uppercase")
    if st["sup"]:
        css.append("vertical-align:super;font-size:.75em")
    if st["sub"]:
        css.append("vertical-align:sub;font-size:.75em")
    return ";".join(css)


def _rtf_new_state():
    return dict(bold=False, italic=False, underline=False, strike=False,
                sup=False, sub=False, caps=False, hidden=False, size=0.0,
                color="", highlight="", align="", li=0, fi=0, sb=0, sa=0)


def rtf_to_html(data):
    r"""RTF -> HTML：中文(cp936 \'xx / \uN)、字符格式、段落对齐、项目符号、PNG/JPEG 图片。"""
    s = data.decode("latin-1")
    n = len(s)
    codec = "cp1252"
    colors = []
    m = re.search(r"\{\\colortbl", s)
    if m:
        st = s.rfind("{", 0, m.start() + 1)
        seg = s[st:_rtf_group_end(s, st)]
        for part in seg.split(";"):
            c = re.search(r"\\red(\d+)\\green(\d+)\\blue(\d+)", part)
            colors.append("#%02x%02x%02x" % tuple(int(c.group(i)) for i in (1, 2, 3))
                          if c else "")
    stack = []
    state = _rtf_new_state()
    paras, frag, pending, bytebuf = [], [], [], bytearray()
    pstyle = [None]
    uc = [1]

    def add_text(t):
        if not t or state["hidden"]:
            return
        if not pending:
            pstyle[0] = dict(state)
        pending.append(t)

    def flush_bytes():
        if bytebuf:
            b = bytes(bytebuf)
            bytebuf.clear()
            try:
                add_text(b.decode(codec, errors="replace"))
            except LookupError:
                add_text(b.decode("cp1252", errors="replace"))

    def flush():
        flush_bytes()
        if not pending:
            return
        txt, css = "".join(pending), _rtf_char_css(pstyle[0] or state)
        frag.append('<span style="%s">%s</span>' % (css, html_lib.escape(txt)) if css
                    else html_lib.escape(txt))
        pending.clear()
        pstyle[0] = None

    def end_para():
        flush()
        body = "".join(frag)
        css = []
        if state["align"]:
            css.append("text-align:" + state["align"])
        if state["li"]:
            css.append("margin-left:%dpx" % state["li"])
        if state["fi"]:
            css.append("text-indent:%dpx" % state["fi"])
        if state["sb"]:
            css.append("margin-top:%dpx" % state["sb"])
        if state["sa"]:
            css.append("margin-bottom:%dpx" % state["sa"])
        paras.append('<p%s>%s</p>' % (' style="%s"' % ";".join(css) if css else "",
                                      body or "&nbsp;"))
        frag.clear()

    def color_at(k):
        try:
            return colors[k] if 0 <= k < len(colors) else ""
        except TypeError:
            return ""

    i = 0
    while i < n:
        c = s[i]
        if c == "{":
            star = s.startswith("\\*", i + 1)   # {\*...} 是可忽略目标，按规范应整体跳过
            word = _rtf_group_word(s, i)
            # 只有 shppict 里包着真正的 \pict 需要进入；其余 {\*...}
            # （ftnsep/ftnsepc/fchars/lchars/cs/generator/datastore...）全是元数据，
            # 之前没跳过会把脚注分隔符、字符集当成正文吐出一串标点垃圾。
            if (star and word != "shppict") or word in RTF_SKIP_DESTS or word == "colortbl":
                i = _rtf_group_end(s, i)
                continue
            if word == "pict":
                end = _rtf_group_end(s, i)
                flush()
                frag.append(_rtf_pict_html(s[i:end]))
                i = end
                continue
            flush()
            stack.append(dict(state))
            i += 1
            continue
        if c == "}":
            flush()
            if stack:
                state = stack.pop()
            i += 1
            continue
        if c == "\\":
            nx = s[i + 1] if i + 1 < n else ""
            if nx == "'":
                try:
                    bytebuf.append(int(s[i + 2:i + 4], 16))
                except ValueError:
                    pass
                i += 4
                continue
            if nx in "\\{}":
                flush_bytes()
                add_text(nx)
                i += 2
                continue
            if nx == "~":
                flush_bytes()
                add_text("\u00a0")
                i += 2
                continue
            if nx in "-_*\r\n":
                i += 2
                continue
            m = re.match(r"[A-Za-z]+", s[i + 1:])
            if not m:
                i += 2
                continue
            word = m.group(0)
            j = i + 1 + len(word)
            pm = re.match(r"-?\d+", s[j:])
            param = int(pm.group(0)) if pm else None
            if pm:
                j += len(pm.group(0))
            if word != "u" and j < n and s[j] == " ":
                j += 1
            i = j
            on = (param != 0) if param is not None else True

            if word == "u" and param is not None:
                flush_bytes()
                if i < n and s[i] == " ":
                    i += 1
                for _ in range(max(0, uc[0])):
                    i = _rtf_unit_end(s, i)
                code = param + 65536 if param < 0 else param
                if 0xE000 <= code <= 0xF8FF:
                    add_text(RTF_WINGDINGS.get(code, "•"))
                elif 0 < code < 0x110000:
                    add_text(chr(code))
                continue
            if word == "uc" and param is not None:
                uc[0] = param
                continue
            if word == "ansicpg" and param is not None:
                codec = RTF_CODECS.get(param, "cp1252")
                continue
            if word in ("par", "page", "column", "sect"):
                if word == "page":
                    end_para()
                    paras.append('<div class="rtf-pagebreak"></div>')
                else:
                    end_para()
                continue
            if word in ("line", "tab", "bullet", "enspace", "emspace", "qmspace",
                        "lquote", "rquote", "ldblquote", "rdblquote", "emdash", "endash"):
                flush()
                frag.append({"line": "<br>", "tab": "&nbsp;&nbsp;&nbsp;&nbsp;",
                             "bullet": "•", "enspace": "&ensp;", "emspace": "&emsp;",
                             "qmspace": "&nbsp;", "lquote": "‘", "rquote": "’",
                             "ldblquote": "“", "rdblquote": "”",
                             "emdash": "—", "endash": "–"}[word])
                continue
            flag = {"b": "bold", "i": "italic", "ul": "underline",
                    "strike": "strike", "super": "sup", "sub": "sub",
                    "scaps": "caps", "caps": "caps", "v": "hidden"}.get(word)
            if flag:
                flush()
                state[flag] = on
                if flag == "sup" and on:
                    state["sub"] = False
                if flag == "sub" and on:
                    state["sup"] = False
                continue
            if word in ("ulnone", "nosupersub"):
                flush()
                if word == "ulnone":
                    state["underline"] = False
                else:
                    state["sup"] = state["sub"] = False
                continue
            if word == "plain":
                flush()
                fresh = _rtf_new_state()
                for k in ("bold", "italic", "underline", "strike", "sup", "sub",
                          "caps", "hidden", "size", "color", "highlight"):
                    state[k] = fresh[k]
                continue
            if word == "pard":
                flush()
                for k in ("align", "li", "fi", "sb", "sa"):
                    state[k] = _rtf_new_state()[k]
                continue
            if word == "fs" and param is not None:
                flush()
                state["size"] = param / 2.0
                continue
            if word in ("cf", "highlight") and param is not None:
                flush()
                state["color" if word == "cf" else "highlight"] = color_at(param)
                continue
            if word == "cb" and param is not None:
                continue
            align = {"qc": "center", "qr": "right", "ql": "left", "qj": "justify"}.get(word)
            if align:
                flush()
                state["align"] = align
                continue
            indent = {"li": "li", "fi": "fi", "sb": "sb", "sa": "sa"}.get(word)
            if indent and param is not None:
                flush()
                state[indent] = max(0, param // 15)
                continue
            continue  # 其余控制字（字体、语言、样式等）忽略，不影响正文
        if c in "\r\n":
            i += 1
            continue
        if ord(c) >= 128:
            bytebuf.append(ord(c))
        else:
            flush_bytes()
            add_text(c)
        i += 1
    flush_bytes()
    if frag or pending:
        end_para()
    return "".join(paras) or '<p>（空文档）</p>'


def rtf_plain_text(data):
    return html_lib.unescape(re.sub(r"<[^>]+>", " ", rtf_to_html(data)))


def walk_files(root, limit=MAX_SEARCH_FILES):
    out = []
    skip_dir = {".git", "node_modules", "__pycache__", ".venv", "venv", "$RECYCLE.BIN", "System Volume Information"}
    try:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in skip_dir and not d.startswith("~$")]
            for n in filenames:
                if n.startswith("~$"):
                    continue
                p = os.path.join(dirpath, n)
                if os.path.splitext(n)[1].lower() in VIEWABLE:
                    out.append(p)
                    if len(out) >= limit:
                        return out
    except OSError:
        pass
    return out


# ---------------------------------------------------------------- Win32 / WPS

if sys.platform == "win32":
    u32 = ctypes.windll.user32
    k32 = ctypes.windll.kernel32
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    GWL_STYLE = -16
    WS_CHILD = 0x40000000
    WS_POPUP = 0x80000000
    WS_CAPTION = 0x00C00000
    WS_THICKFRAME = 0x00040000
    WS_SYSMENU = 0x00080000
    WS_MINIMIZEBOX = 0x00020000
    WS_MAXIMIZEBOX = 0x00010000
    WS_VISIBLE = 0x10000000
    HWND_TOP = 0
    SWP_NOZORDER = 0x0004
    SWP_FRAMECHANGED = 0x0020

    # 64 位下必须为 P/Invoke 声明签名，否则句柄/样式会被截成 32 位 int
    HWND_T = wintypes.HWND
    RECT_P = ctypes.POINTER(wintypes.RECT)
    u32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
    u32.EnumWindows.restype = wintypes.BOOL
    u32.IsWindowVisible.argtypes = [HWND_T]
    u32.GetWindowTextW.argtypes = [HWND_T, wintypes.LPWSTR, ctypes.c_int]
    u32.GetClassNameW.argtypes = [HWND_T, wintypes.LPWSTR, ctypes.c_int]
    u32.GetWindowRect.argtypes = [HWND_T, RECT_P]
    u32.GetClientRect.argtypes = [HWND_T, RECT_P]
    u32.GetWindowThreadProcessId.argtypes = [HWND_T, ctypes.POINTER(wintypes.DWORD)]
    u32.MoveWindow.argtypes = [HWND_T] + [ctypes.c_int] * 4 + [wintypes.BOOL]
    u32.SetWindowPos.argtypes = [HWND_T, HWND_T] + [ctypes.c_int] * 4 + [wintypes.UINT]
    u32.SetParent.argtypes = [HWND_T, HWND_T]
    u32.SetParent.restype = HWND_T
    u32.GetParent.argtypes = [HWND_T]
    u32.GetParent.restype = HWND_T
    u32.ShowWindow.argtypes = [HWND_T, ctypes.c_int]
    for fn in ("GetWindowLongPtrW", "SetWindowLongPtrW", "GetWindowLongW", "SetWindowLongW"):
        if hasattr(u32, fn):
            f = getattr(u32, fn)
            if fn.startswith("Get"):
                f.argtypes, f.restype = [HWND_T, ctypes.c_int], ctypes.c_ssize_t
            else:
                f.argtypes, f.restype = [HWND_T, ctypes.c_int, ctypes.c_ssize_t], ctypes.c_ssize_t
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                               wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]

    try:  # 让宿主与浏览器的 DPI 认知一致，减少高 DPI 下的坐标漂移
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            u32.SetProcessDPIAware()
        except Exception:
            pass

    def get_style(h):
        f = getattr(u32, "GetWindowLongPtrW", None) or u32.GetWindowLongW
        return f(h, GWL_STYLE)

    def set_style(h, s):
        f = getattr(u32, "SetWindowLongPtrW", None) or u32.SetWindowLongW
        return f(h, GWL_STYLE, s)

    def enum_windows():
        found = []

        def cb(hwnd, _):
            if u32.IsWindowVisible(hwnd):
                buf = ctypes.create_unicode_buffer(512)
                u32.GetWindowTextW(hwnd, buf, 512)
                cls = ctypes.create_unicode_buffer(256)
                u32.GetClassNameW(hwnd, cls, 256)
                found.append((hwnd, buf.value, cls.value))
            return True

        u32.EnumWindows(WNDENUMPROC(cb), 0)
        return found

    def window_rect(hwnd):
        r = wintypes.RECT()
        u32.GetWindowRect(hwnd, ctypes.byref(r))
        return (r.left, r.top, r.right - r.left, r.bottom - r.top)

    def client_rect(hwnd):
        r = wintypes.RECT()
        u32.GetClientRect(hwnd, ctypes.byref(r))
        return (r.right - r.left, r.bottom - r.top)

    def proc_name(hwnd):
        pid = wintypes.DWORD()
        u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        h = k32.OpenProcess(0x1000, False, pid.value)  # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = wintypes.DWORD(1024)
            if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                return os.path.basename(buf.value).lower()
            return ""
        finally:
            k32.CloseHandle(h)

    # ---- 窗口识别：新窗口标题固定为 UniView，不再依赖露在标题栏上的随机码 ----
    def find_uniview_window():
        """最小化窗口的 WS_VISIBLE 仍为真，所以同样能被枚举到。"""
        legacy = None
        for hwnd, title, cls in enum_windows():
            t = (title or "").strip()
            clean = (t == "UniView")
            if not (clean or TITLE_TAG in t):
                continue
            if proc_name(hwnd) not in ("msedge.exe", "chrome.exe"):
                continue
            if clean:
                return hwnd          # 精确匹配优先，避免命中旧版本窗口
            legacy = legacy or hwnd
        return legacy

    # ---- 强制置前：从「打开方式」打开文件时把最小化的窗口还原并抢到最前 ----
    u32.IsIconic.argtypes = [HWND_T]
    u32.IsIconic.restype = wintypes.BOOL
    u32.SetForegroundWindow.argtypes = [HWND_T]
    u32.SetForegroundWindow.restype = wintypes.BOOL
    u32.BringWindowToTop.argtypes = [HWND_T]
    u32.BringWindowToTop.restype = wintypes.BOOL
    u32.GetForegroundWindow.restype = HWND_T
    u32.FlashWindow.argtypes = [HWND_T, wintypes.BOOL]
    u32.FlashWindow.restype = wintypes.BOOL
    u32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
    u32.AttachThreadInput.restype = wintypes.BOOL
    u32.keybd_event.argtypes = [ctypes.c_ubyte, ctypes.c_ubyte,
                                wintypes.DWORD, ctypes.c_void_p]
    k32.GetCurrentThreadId.restype = wintypes.DWORD

    def focus_window():
        try:
            h = find_uniview_window()
            if not h:
                log("FOCUS | 未找到 UniView 窗口")
                return False
            if u32.IsIconic(h):
                u32.ShowWindow(h, 9)                 # SW_RESTORE
            # 模拟一次 ALT，解除 Windows 的前台锁定
            u32.keybd_event(0x12, 0, 0, None)
            u32.keybd_event(0x12, 0, 2, None)       # KEYEVENTF_KEYUP
            u32.BringWindowToTop(h)
            ok = bool(u32.SetForegroundWindow(h))
            if not ok:
                fg = u32.GetForegroundWindow()
                tid_fg = u32.GetWindowThreadProcessId(fg, None)
                tid_me = k32.GetCurrentThreadId()
                if tid_fg and tid_fg != tid_me:
                    u32.AttachThreadInput(tid_me, tid_fg, True)
                    u32.SetForegroundWindow(h)
                    u32.BringWindowToTop(h)
                    u32.AttachThreadInput(tid_me, tid_fg, False)
                    ok = bool(u32.GetForegroundWindow() == h)
            if not ok:
                u32.FlashWindow(h, True)             # 实在抢不到，至少任务栏闪一下
            log("FOCUS | restore+foreground=%s hwnd=%s" % (ok, h))
            return ok
        except Exception as ex:
            log("FOCUS error | " + str(ex))
            return False

    # ---------------------------------------------------------------- 系统托盘
    # 关闭窗口后服务不退出，驻留托盘；双击托盘图标/再次双击桌面图标重开窗口，
    # 托盘右键菜单可「真正退出」。纯标准库：隐藏消息窗口 + Shell_NotifyIconW。
    shell32 = ctypes.windll.shell32

    WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND,
                                 wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

    class WNDCLASSW(ctypes.Structure):
        _fields_ = [("style", wintypes.UINT), ("lpfnWndProc", WNDPROC),
                    ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int),
                    ("hInstance", wintypes.HINSTANCE), ("hIcon", wintypes.HICON),
                    ("hCursor", wintypes.HANDLE), ("hbrBackground", wintypes.HBRUSH),
                    ("lpszMenuName", wintypes.LPCWSTR), ("lpszClassName", wintypes.LPCWSTR)]

    class NOTIFYICONDATAW(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("hWnd", wintypes.HWND),
                    ("uID", wintypes.UINT), ("uFlags", wintypes.UINT),
                    ("uCallbackMessage", wintypes.UINT), ("hIcon", wintypes.HICON),
                    ("szTip", wintypes.WCHAR * 128),
                    ("dwState", wintypes.DWORD), ("dwStateMask", wintypes.DWORD),
                    ("szInfo", wintypes.WCHAR * 256),
                    ("uTimeoutOrVersion", wintypes.UINT),
                    ("szInfoTitle", wintypes.WCHAR * 64),
                    ("dwInfoFlags", wintypes.DWORD),
                    ("guidItem", ctypes.c_byte * 16),
                    ("hBalloonIcon", wintypes.HICON)]

    u32.RegisterWindowMessageW.argtypes = [wintypes.LPCWSTR]
    u32.RegisterWindowMessageW.restype = wintypes.UINT
    u32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                                    wintypes.DWORD] + [ctypes.c_int] * 4 + [
                                    wintypes.HWND, wintypes.HMENU,
                                    wintypes.HINSTANCE, wintypes.LPVOID]
    u32.CreateWindowExW.restype = wintypes.HWND
    u32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT,
                                   wintypes.WPARAM, wintypes.LPARAM]
    u32.DefWindowProcW.restype = ctypes.c_ssize_t
    u32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT,
                               ctypes.c_int, ctypes.c_int, wintypes.UINT]
    u32.LoadImageW.restype = wintypes.HICON
    u32.CreatePopupMenu.restype = wintypes.HMENU
    u32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_size_t,
                                wintypes.LPCWSTR]
    u32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
    u32.TrackPopupMenu.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_int,
                                   ctypes.c_int, ctypes.c_int, wintypes.HWND,
                                   wintypes.LPVOID]
    u32.DestroyMenu.argtypes = [wintypes.HMENU]
    u32.PostQuitMessage.argtypes = [ctypes.c_int]
    u32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND,
                                wintypes.UINT, wintypes.UINT]
    u32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
    u32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
    u32.DestroyWindow.argtypes = [wintypes.HWND]
    k32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    k32.GetModuleHandleW.restype = wintypes.HINSTANCE
    shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD,
                                          ctypes.POINTER(NOTIFYICONDATAW)]

    class TrayIcon:
        WM_TRAY = 0x0401          # WM_USER + 1
        WM_COMMAND = 0x0111
        CMD_OPEN = 1001
        CMD_QUIT = 1002
        NIM_ADD, NIM_MODIFY, NIM_DELETE, NIM_SETVERSION = 0, 1, 2, 4
        NIF_MESSAGE, NIF_ICON, NIF_TIP, NIF_INFO = 0x1, 0x2, 0x4, 0x10

        def __init__(self, icon_path):
            self.icon_path = icon_path
            self.hwnd = None
            self.hicon = None
            self.added = False
            self._notified = False
            self._thread = None

        def start(self):
            if self._thread is None:
                self._thread = threading.Thread(target=self._run, daemon=True)
                self._thread.start()

        def _data(self, flags):
            d = NOTIFYICONDATAW()
            d.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
            d.hWnd = self.hwnd
            d.uID = 1
            d.uFlags = flags
            if flags & self.NIF_MESSAGE:
                d.uCallbackMessage = self.WM_TRAY
            if flags & self.NIF_ICON:
                d.hIcon = self.hicon
            if flags & self.NIF_TIP:
                d.szTip = "UniView 文件查看器"
            return d

        def _notify(self, msg, flags):
            if self.hwnd:
                shell32.Shell_NotifyIconW(msg, ctypes.byref(self._data(flags)))

        def _run(self):
            icon_flags = self.NIF_MESSAGE | self.NIF_ICON | self.NIF_TIP
            self._taskbar_created = u32.RegisterWindowMessageW("TaskbarCreated")
            self._wndproc = WNDPROC(self._on_msg)     # 必须持引用，防被 GC
            wc = WNDCLASSW()
            wc.lpfnWndProc = self._wndproc
            wc.hInstance = k32.GetModuleHandleW(None)
            wc.lpszClassName = "UniViewTrayWnd"
            if not u32.RegisterClassW(ctypes.byref(wc)):
                err = k32.GetLastError()
                if err != 1410:                        # 1410 = 类已注册，可复用
                    log("TRAY RegisterClass failed | " + str(err))
                    return
            self.hwnd = u32.CreateWindowExW(0, "UniViewTrayWnd", "UniViewTray",
                                            0, 0, 0, 0, 0, wintypes.HWND(-3),
                                            None, wc.hInstance, None)
            if not self.hwnd:
                log("TRAY CreateWindow failed | " + str(k32.GetLastError()))
                return
            self.hicon = u32.LoadImageW(None, self.icon_path, 1, 0, 0, 0x10 | 0x40)
            # 1=IMAGE_ICON, 0x10=LR_LOADFROMFILE, 0x40=LR_DEFAULTSIZE
            if not self.hicon:
                self.hicon = u32.LoadIconW(None, 32512)   # 兜底：IDI_APPLICATION
            self._notify(self.NIM_ADD, icon_flags)
            self._notify(self.NIM_SETVERSION, 0)
            self.added = True
            log("TRAY ready")
            msg = wintypes.MSG()
            while u32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                u32.TranslateMessage(ctypes.byref(msg))
                u32.DispatchMessageW(ctypes.byref(msg))
            self.added = False

        def _on_msg(self, hwnd, msg, wp, lp):
            if msg == self.WM_TRAY:
                if lp in (0x0202, 0x0203):             # 左键单击/双击：重开窗口
                    ensure_window_open()
                elif lp == 0x0205:                     # 右键：弹菜单
                    self._menu()
                return 0
            if msg == self.WM_COMMAND:
                cmd = wp & 0xFFFF
                if cmd == self.CMD_OPEN:
                    ensure_window_open()
                elif cmd == self.CMD_QUIT:
                    shutdown_app()
                return 0
            if msg == self._taskbar_created and self.added:
                # 资源管理器重启后托盘区清空，需要重新挂图标
                self._notify(self.NIM_ADD,
                             self.NIF_MESSAGE | self.NIF_ICON | self.NIF_TIP)
                return 0
            return u32.DefWindowProcW(hwnd, msg, wp, lp)

        def _menu(self):
            menu = u32.CreatePopupMenu()
            if not menu:
                return
            u32.AppendMenuW(menu, 0, self.CMD_OPEN, "打开 UniView")
            u32.AppendMenuW(menu, 0x800, 0, None)      # MF_SEPARATOR
            u32.AppendMenuW(menu, 0, self.CMD_QUIT, "退出 UniView")
            u32.SetForegroundWindow(self.hwnd)         # 菜单失焦自动收起的标准做法
            pt = wintypes.POINT()
            u32.GetCursorPos(ctypes.byref(pt))
            u32.TrackPopupMenu(menu, 0x0002, pt.x, pt.y, 0, self.hwnd, None)
            u32.PostMessageW(self.hwnd, 0, 0, 0)
            u32.DestroyMenu(menu)

        def notify_once(self, title, text):
            """关窗转入托盘时提示一次；不是每次关窗都弹。"""
            if self._notified or not self.added:
                return
            self._notified = True
            d = self._data(self.NIF_INFO)
            d.szInfo = text[:255]
            d.szInfoTitle = title[:63]
            d.dwInfoFlags = 1                          # NIIF_INFO
            shell32.Shell_NotifyIconW(self.NIM_MODIFY, ctypes.byref(d))

        def remove(self):
            if self.added:
                self._notify(self.NIM_DELETE, 0)
                self.added = False

else:
    u32 = None
    focus_window = None

    class TrayIcon:        # 非 Windows 平台占位：不起托盘
        def __init__(self, icon_path):
            pass

        def start(self):
            pass

        def notify_once(self, title, text):
            pass

        def remove(self):
            pass


# 托盘实例：图标文件从资源目录读（打包后是 _MEIPASS 里随包带的 uniview.ico）
TRAY = TrayIcon(os.path.join(ROOT, "uniview.ico"))


FOCUS_LOCK = threading.Lock()
LAST_FOCUS = 0.0


def raise_window():
    """节流后异步置前，避免两个文件同时右键时重复抢焦点。"""
    global LAST_FOCUS
    if not callable(focus_window):
        return
    if BROWSE_ACTIVE.is_set():   # 系统文件夹对话框开着时，别把主窗口抢回最前
        return
    with FOCUS_LOCK:
        now = time.time()
        if now - LAST_FOCUS < 1.5:
            return
        LAST_FOCUS = now
    threading.Thread(target=focus_window, daemon=True).start()


def find_wps_exe(ext):
    """按扩展名挑 WPS 组件；找不到就返回 None（降级到系统默认程序）。"""
    base_variants = [
        os.path.expandvars(r"%LOCALAPPDATA%\Kingsoft\WPS Office"),
        r"C:\Program Files (x86)\Kingsoft\WPS Office",
        r"C:\Program Files\Kingsoft\WPS Office",
    ]
    want = "wps.exe"
    if ext in {".xls", ".xlsx", ".et", ".csv"}:
        want = "et.exe"
    elif ext in {".ppt", ".pptx", ".dps"}:
        want = "wpp.exe"
    hits = []
    for base in base_variants:
        hits.extend(glob.glob(os.path.join(base, "*", "office6", want)))
    if not hits:
        # 任意 WPS 主程序都能打开常见格式，退而求其次
        for base in base_variants:
            hits.extend(glob.glob(os.path.join(base, "*", "office6", "wps.exe")))
    if not hits:
        return None
    hits.sort()
    return hits[-1]


class OfficeHost:
    """把 WPS 真实窗口 SetParent 到浏览器窗口内的一块区域。

    这是「硬嵌」路径：官方 COM/OLE 嵌入需要企业版，个人版只能走这条。
    探针的第 3/4 项（IME 候选框、缩放跟随）决定它能否日常使用，
    因此这里做成「尝试 + 一键降级」，绝不把用户文件锁死在这条路上。
    """

    def __init__(self):
        self.host_hwnd = None
        self.wps_hwnd = None
        self.old_style = None
        self.old_rect = None
        self.rect = None            # css 像素，相对浏览器 viewport
        self.scale = 1.0
        self.lock = threading.Lock()
        self.embed_lock = threading.Lock()
        self.status = "idle"
        self.message = ""
        self.diag = []
        self._stop = threading.Event()
        self._follower = threading.Thread(target=self._follow, daemon=True)

    # ---- 找浏览器窗口：标题固定为 UniView，按标题 + 进程名精确识别
    def wait_host(self, timeout=18):
        deadline = time.time() + timeout
        while time.time() < deadline and not self._stop.is_set():
            h = find_uniview_window()
            if h:
                self.host_hwnd = h
                return True
            time.sleep(0.25)
        return False

    def start_follower(self):
        if not self._follower.is_alive():
            self._follower.start()

    def _follow(self):
        while not self._stop.wait(0.35):
            with self.lock:
                if self.wps_hwnd and self.rect:
                    try:
                        self._place(self.rect, self.scale)
                    except Exception:
                        pass

    def _place(self, rect, scale):
        if not (self.host_hwnd and self.wps_hwnd):
            return
        cw, ch = client_rect(self.host_hwnd)
        x = int(rect["x"] * scale)
        y = int(rect["y"] * scale)
        w = max(120, min(int(rect["w"] * scale), cw - x + 4))
        h = max(120, min(int(rect["h"] * scale), ch - y + 4))
        if x > cw or y > ch:
            return
        u32.MoveWindow(self.wps_hwnd, x, y, w, h, True)
        u32.SetWindowPos(self.wps_hwnd, HWND_TOP, 0, 0, 0, 0,
                         SWP_NOZORDER | 0x0001 | 0x0010)  # NOSIZE|NOACTIVATE

    def sync(self, rect, scale=1.0):
        with self.lock:
            self.rect = rect
            self.scale = float(scale) if 0.5 <= float(scale) <= 4 else 1.0
            if self.wps_hwnd:
                self._place(self.rect, self.scale)

    def embed(self, path, rect, scale=1.0):
        # 多个标签快速切换时，保证 embedding 不并发互相踩
        with self.embed_lock:
            return self._embed_impl(path, rect, scale)

    def _embed_impl(self, path, rect, scale=1.0):
        if sys.platform != "win32":
            self.status, self.message = "unsupported", "仅 Windows 支持窗口嵌入"
            return self.info()
        if self.wps_hwnd:
            self.detach()

        if not self.host_hwnd and not self.wait_host(8):
            self.status, self.message = "fail", "没找到软件主窗口，无法嵌入"
            return self.info()

        exe = find_wps_exe(os.path.splitext(path)[1].lower())
        if not exe:
            self.status, self.message = "no_wps", "没找到 WPS 可执行文件"
            return self.info()

        # 记录启动前的窗口标题：WPS 多标签复用主窗口时 HWND 不变，只有标题会变，
        # 只比对 HWND 集合会永远等不到「新窗口」。
        before = {h: t for h, t, _ in enum_windows()}
        name = os.path.basename(path)
        stem = os.path.splitext(name)[0]
        log("EMBED start | file=%s | exe=%s | host=%s" % (name, exe, self.host_hwnd))
        try:
            subprocess.Popen([exe, path], close_fds=True)
        except Exception as ex:
            self.status, self.message = "fail", "启动 WPS 失败：" + str(ex)
            log("EMBED launch exception " + repr(ex))
            return self.info()

        target = None
        reused = False
        fresh = set()
        diag = []
        deadline = time.time() + 10
        while time.time() < deadline:
            for hwnd, title, cls in enum_windows():
                if hwnd == self.host_hwnd or not title.strip():
                    continue
                pn = proc_name(hwnd)
                looks_wps = ("wps" in pn or pn in ("et.exe", "wpp.exe")
                             or "kingsoft" in title.lower() or "wps" in title.lower())
                is_new = hwnd not in before
                changed = (not is_new) and before.get(hwnd) != title
                if is_new:
                    fresh.add(hwnd)
                if looks_wps or is_new:
                    s = "[%s] %s / %s%s" % (pn or "?", title[:60], cls[:24],
                                            " NEW" if is_new else (" CHANGED" if changed else ""))
                    if s not in diag:
                        diag.append(s)
                # 文档窗口的标题几乎一定含文件名主干（含或不含扩展名）
                if stem and stem in title:
                    target, reused = hwnd, not is_new
                    break
                if is_new and looks_wps and re.search(r"\.(docx?|xlsx?|pptx?|pdf)\b", title, re.I):
                    target = hwnd
                    break
            if target:
                break
            time.sleep(0.35)
        if not target:
            # 放宽一轮：进程名可能不是 wps.exe（云进程/兼容层），只认标题
            for hwnd, title, cls in enum_windows():
                if hwnd != self.host_hwnd and stem and stem in title:
                    target, reused = hwnd, True
                    break
        self.diag = diag[-14:]
        log("EMBED windows | " + " ; ".join(self.diag))

        if not target:
            # 不关闭 WPS（可能已有未保存内容），只把新弹出来的窗口收起来，别挡住软件
            for h in fresh:
                try:
                    u32.ShowWindow(h, 6)  # SW_MINIMIZE
                except Exception:
                    pass
            self.status = "fail"
            self.message = "没识别到 WPS 文档窗口（新弹出的已最小化）；详情见 uniview.log，可点「独立窗口打开」"
            log("EMBED FAIL no target")
            return self.info()

        try:
            self.old_rect = window_rect(target)
            self.old_style = get_style(target)
            u32.SetParent(target, self.host_hwnd)
            parent = u32.GetParent(target)
            if int(parent or 0) != int(self.host_hwnd):
                raise RuntimeError("SetParent 后父窗口不匹配 (parent=%s host=%s)" % (parent, self.host_hwnd))
            style = (self.old_style | WS_CHILD | WS_VISIBLE) & ~(
                WS_POPUP | WS_CAPTION | WS_THICKFRAME | WS_SYSMENU |
                WS_MINIMIZEBOX | WS_MAXIMIZEBOX)
            set_style(target, style)
            u32.SetWindowPos(target, 0, 0, 0, 0, 0,
                             SWP_NOZORDER | 0x0001 | 0x0002 | SWP_FRAMECHANGED)
            self.wps_hwnd = target
            self.sync(rect, scale)
            self.status = "embedded"
            self.message = ("WPS 已嵌入（复用了已有主窗口，其他标签页会一起带进来）；"
                            "在文档里改字由 WPS 自己保存，原文件零损耗" if reused
                            else "WPS 已嵌入；在文档里改字由 WPS 自己保存，原文件零损耗")
            self.start_follower()
            log("EMBED OK hwnd=%s reused=%s" % (target, reused))
        except Exception as ex:
            log("EMBED EXC " + repr(ex))
            try:
                u32.SetParent(target, 0)
                if self.old_style is not None:
                    set_style(target, self.old_style)
                u32.ShowWindow(target, 6)
            except Exception:
                pass
            self.status, self.message = "fail", "嵌入失败：" + str(ex) + "（WPS 窗口已最小化）"
        return self.info()

    def detach(self):
        with self.lock:
            if not self.wps_hwnd:
                self.status, self.message = "idle", ""
                return self.info()
            hwnd = self.wps_hwnd
            try:
                u32.SetParent(hwnd, 0)
                if self.old_style is not None:
                    set_style(hwnd, self.old_style)
                if self.old_rect:
                    x, y, w, h = self.old_rect
                    u32.MoveWindow(hwnd, x, y, max(600, w), max(400, h), True)
                u32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                 SWP_NOZORDER | 0x0001 | 0x0002 | SWP_FRAMECHANGED)
            except Exception:
                pass
            self.wps_hwnd = None
            self.status = "detached"
            self.message = "已还原为 WPS 独立窗口，未保存的编辑仍在 WPS 里，不会丢失"
        return self.info()

    def info(self):
        return {"status": self.status, "message": self.message,
                "host": bool(self.host_hwnd), "wps": find_wps_exe(".docx") is not None,
                "diag": self.diag}


HOST = OfficeHost()


# ---------------------------------------------------------------- HTTP 服务

class Server(ThreadingHTTPServer):
    daemon_threads = True

    def handle_error(self, request, client_address):
        # 浏览器刷新/切换标签会中止请求，http.server 默认会刷一屏 ConnectionAbortedError
        exc = sys.exc_info()[1]
        if isinstance(exc, (ConnectionError, BrokenPipeError)):
            return
        super().handle_error(request, client_address)


class Handler(BaseHTTPRequestHandler):
    server_version = "UniView/0.1"

    def log_message(self, fmt, *args):
        pass  # 自用，不刷屏

    def _auth(self, q):
        return q.get("t", [""])[0] == TOKEN

    def _json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _raw(self, data, ctype, code=200, headers=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        n = int(self.headers.get("Content-Length", 0) or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        u = urllib.parse.urlsplit(self.path)
        q = urllib.parse.parse_qs(u.query)
        try:
            if u.path in ("/", "/index.html"):
                with open(os.path.join(ROOT, "index.html"), "rb") as f:
                    return self._raw(f.read(), "text/html; charset=utf-8")
            if u.path.startswith("/vendor/"):
                name = os.path.basename(u.path)
                if name not in ("jszip.min.js", "xlsx.full.min.js", "marked.min.js", "mermaid.min.js"):
                    return self._json({"error": "forbidden"}, 403)
                with open(os.path.join(ROOT, "vendor", name), "rb") as f:
                    return self._raw(f.read(), "application/javascript; charset=utf-8")

            if u.path == "/favicon.ico":
                p = os.path.join(ROOT, "uniview.ico")
                if not os.path.isfile(p):
                    return self._json({"error": "no icon"}, 404)
                with open(p, "rb") as f:
                    return self._raw(f.read(), "image/x-icon")

            if u.path.startswith("/html/"):
                # HTML 隔离渲染：主文档 + 同目录资源代理。钥匙在路径里，
                # 与 API 令牌分开，沙箱页面即使读到 URL 也拿不到真正的令牌。
                parts = u.path.split("/")
                if len(parts) < 4 or parts[2] != HTML_KEY:
                    return self._json({"error": "forbidden"}, 403)
                try:
                    doc = _b64d(parts[3])
                except Exception:
                    return self._json({"error": "bad path"}, 400)
                if not os.path.isfile(doc):
                    return self._json({"error": "not a file"}, 404)
                home_dir = os.path.abspath(os.path.dirname(doc))
                allow = os.path.dirname(home_dir) or home_dir
                rel = urllib.parse.unquote("/".join(parts[4:]))
                if not rel:
                    origin = "http://127.0.0.1:%d" % PORT
                    csp = ("default-src {0} data: blob:; "
                           "script-src {0} 'unsafe-inline' 'unsafe-eval' data: blob:; "
                           "style-src {0} 'unsafe-inline' data: blob:; "
                           "img-src {0} data: blob:; font-src {0} data: blob:; "
                           "media-src {0} data: blob:; connect-src 'none'; "
                           "form-action 'none'; object-src 'none'; frame-src 'none'; "
                           "base-uri {0}").format(origin)
                    base_url = "/html/%s/%s" % (HTML_KEY, parts[3])
                    return self._raw(build_html_document(doc, base_url),
                                     "text/html; charset=utf-8",
                                     headers={"Content-Security-Policy": csp})
                cand = os.path.abspath(os.path.normpath(os.path.join(home_dir, rel)))
                try:
                    if os.path.commonpath([cand, allow]) != allow:
                        return self._json({"error": "forbidden"}, 403)
                except ValueError:
                    return self._json({"error": "forbidden"}, 403)
                if not os.path.isfile(cand):
                    return self._json({"error": "not a file"}, 404)
                with open(cand, "rb") as f:
                    data = f.read()
                ctype = mimetypes.guess_type(cand)[0] or "application/octet-stream"
                if os.path.splitext(cand)[1].lower() in TEXT_EXT:
                    ctype += "; charset=utf-8"
                return self._raw(data, ctype)

            if not self._auth(q):
                return self._json({"error": "unauthorized"}, 401)

            if u.path == "/api/ping":
                return self._json({"ok": True, "pid": os.getpid(), "port": PORT})

            if u.path == "/api/pending":
                # 前端每次轮询都证明窗口还活着；一次性取走所有待打开文件
                STATE["seen"] = time.time()
                with PENDING_LOCK:
                    files = PENDING[:]
                    PENDING.clear()
                if files:
                    raise_window()   # 窗口可能被最小化：取到文件时也把它弹起来
                return self._json({"files": files})

            if u.path == "/api/boot":
                home = os.path.expanduser("~")
                samples = os.path.normpath(os.path.join(ROOT, "..", "web-viewer", "samples"))
                roots = []
                # 打包后没有随身的示例目录，就不显示「示例文件」入口，避免指到解包临时目录
                if os.path.isdir(samples):
                    roots.append({"name": "示例文件", "path": samples})
                for n, p in [("桌面", os.path.join(home, "Desktop")),
                             ("文档", os.path.join(home, "Documents")),
                             ("下载", os.path.join(home, "Downloads")),
                             ("主目录", home)]:
                    if os.path.isdir(p):
                        roots.append({"name": n, "path": p})
                roots.append({"name": "此电脑", "path": THIS_PC})
                return self._json({"roots": roots, "customRoots": read_custom_roots(),
                                   "titleTag": TITLE_TAG,
                                   "htmlKey": HTML_KEY,
                                   "wps": find_wps_exe(".docx") is not None,
                                   "office": HOST.info(), "token": TOKEN})

            if u.path == "/api/session":
                # 上次退出/关窗前打开过的标签；文件已不存在的一律不返回
                return self._json({"session": read_session()})

            if u.path == "/api/tree":
                p = q.get("path", [""])[0]
                if p == THIS_PC:
                    return self._json({"items": list_drives(), "hidden": 0})
                if q.get("probe", ["0"])[0] == "1":
                    return self._json({"exists": os.path.isdir(p)})
                if not os.path.isdir(p):
                    return self._json({"error": "not a directory"}, 404)
                # 默认只列出「软件能打开」的文件；all=1 时列出全部（界面开关用）
                show_all = q.get("all", ["0"])[0] == "1"
                items = []
                hidden = 0
                try:
                    for name in os.listdir(p):
                        full = os.path.join(p, name)
                        if name.startswith("~$"):
                            continue
                        try:
                            isd = os.path.isdir(full)
                        except OSError:
                            continue
                        ext = "" if isd else os.path.splitext(name)[1].lower()
                        viewable = isd or ext in VIEWABLE
                        if not show_all and not viewable:
                            hidden += 1
                            continue
                        try:
                            size = 0 if isd else os.path.getsize(full)
                        except OSError:
                            size = 0
                        items.append({"name": name, "path": full, "dir": isd,
                                      "ext": ext, "viewable": viewable,
                                      "size": size})
                except OSError as ex:
                    return self._json({"error": str(ex)}, 403)
                items.sort(key=lambda x: (not x["dir"], x["name"].lower()))
                return self._json({"items": items, "hidden": hidden, "all": show_all})

            if u.path == "/api/text":
                p = q.get("path", [""])[0]
                if not os.path.isfile(p):
                    return self._json({"error": "not a file"}, 404)
                text, enc = read_text(p, limit=None)
                return self._json({"text": text, "encoding": enc,
                                   "kind": kind_of(p), "size": os.path.getsize(p)})

            if u.path == "/api/rtf":
                p = q.get("path", [""])[0]
                if not os.path.isfile(p):
                    return self._json({"error": "not a file"}, 404)
                with open(p, "rb") as f:
                    data = f.read()
                try:
                    return self._json({"html": rtf_to_html(data), "size": len(data)})
                except Exception as ex:
                    return self._json({"error": "RTF 解析失败：" + str(ex)}, 500)

            if u.path == "/api/pptx":
                p = q.get("path", [""])[0]
                if not os.path.isfile(p):
                    return self._json({"error": "not a file"}, 404)
                try:
                    return self._json(build_pptx_preview(p))
                except Exception as ex:
                    return self._json({"error": "PPT 解析失败：" + str(ex)}, 500)

            if u.path == "/api/office/media":
                # 只按 zip 内的媒体名取单个文件，不把 297MB 的整包送给浏览器
                p = q.get("path", [""])[0]
                name = q.get("name", [""])[0]
                if not os.path.isfile(p):
                    return self._json({"error": "not a file"}, 404)
                if not re.match(r"^ppt/media/[A-Za-z0-9_.\-]+$", name):
                    return self._json({"error": "forbidden"}, 403)
                try:
                    import zipfile
                    with zipfile.ZipFile(p) as z:
                        data = z.read(name)
                except Exception as ex:
                    return self._json({"error": str(ex)}, 404)
                ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
                if name.lower().endswith(".svg"):
                    ctype = "image/svg+xml"
                return self._raw(data, ctype)

            if u.path == "/api/raw":
                p = q.get("path", [""])[0]
                if not os.path.isfile(p):
                    return self._json({"error": "not a file"}, 404)
                with open(p, "rb") as f:
                    data = f.read()
                ext = os.path.splitext(p)[1].lower()
                ctype = mimetypes.guess_type(p)[0] or "application/octet-stream"
                if ext in {".svg"}:
                    ctype = "image/svg+xml"
                if ctype == "application/octet-stream":
                    ctype = sniff_image_type(data) or ctype
                if ext in TEXT_EXT:
                    ctype += "; charset=utf-8" if "charset" not in ctype else ""
                return self._raw(data, ctype)

            if u.path == "/api/search":
                term = q.get("q", [""])[0].strip()
                root = q.get("root", [""])[0]
                if len(term) < 1 or not os.path.isdir(root):
                    return self._json({"hits": [], "scanned": 0})
                hits, scanned = [], 0
                low = term.lower()
                for p in walk_files(root):
                    try:
                        if os.path.getsize(p) > MAX_SEARCH_BYTES:
                            continue
                    except OSError:
                        continue
                    scanned += 1
                    txt = extract_text(p)
                    if not txt:
                        continue
                    for m in re.finditer(re.escape(low), txt.lower()):
                        a = max(0, m.start() - 45)
                        b = min(len(txt), m.end() + 65)
                        snippet = txt[a:b].replace("\n", " ").replace("\r", " ")
                        hits.append({"path": p, "name": os.path.basename(p),
                                     "snippet": snippet})
                        break
                    if len(hits) >= 60:
                        break
                return self._json({"hits": hits, "scanned": scanned})

            if u.path == "/api/office":
                return self._json(HOST.info())

            return self._json({"error": "not found"}, 404)
        except Exception as ex:
            traceback.print_exc()
            return self._json({"error": str(ex)}, 500)

    def do_POST(self):
        u = urllib.parse.urlsplit(self.path)
        q = urllib.parse.parse_qs(u.query)
        try:
            if not self._auth(q):
                return self._json({"error": "unauthorized"}, 401)
            body = self._body()

            if u.path == "/api/queue":
                p = body.get("path", "")
                if not p or not os.path.isfile(p):
                    return self._json({"error": "bad request"}, 400)
                with PENDING_LOCK:
                    PENDING.append(os.path.abspath(p))
                # 窗口已关闭 / 长时间没轮询时，仍要能从「打开方式」唤回一个窗口。
                # 90 秒阈值：浏览器最小化后会把定时器节流到每分钟一次，
                # 阈值太短会把「最小化」误判成「窗口已关」而多开一个窗口。
                if STATE["seen"] <= 0 or time.time() - STATE["seen"] > 90:
                    url = "http://127.0.0.1:%d/?t=%s" % (PORT, TOKEN)
                    threading.Thread(target=open_window, args=(url,), daemon=True).start()
                    log("QUEUE reopen window | " + p)
                raise_window()       # 最小化时还原并抢到最前
                log("QUEUE file | " + p)
                return self._json({"ok": True, "queued": len(PENDING)})

            if u.path == "/api/window-gone":
                # 窗口关闭时前端发 beacon；服务转入托盘驻留，下次「打开方式」或
                # 双击图标都会重新开窗。第一次转托盘时给个气泡提示，避免用户以为软件没了。
                STATE["seen"] = 0.0
                log("WINDOW gone | 转入托盘驻留")
                TRAY.notify_once("UniView 仍在运行",
                                 "窗口已收起，服务驻留在系统托盘。\n双击托盘图标或桌面图标可重新打开。")
                return self._json({"ok": True})

            if u.path == "/api/session":
                write_session(body)
                return self._json({"ok": True})

            if u.path == "/api/reopen":
                # 托盘菜单/再次双击图标：窗口还在就拉到最前，不在就重开一个
                threading.Thread(target=ensure_window_open, daemon=True).start()
                return self._json({"ok": True})

            if u.path == "/api/roots/add":
                p = body.get("path", "")
                if not p or not os.path.isdir(p):
                    return self._json({"error": "not a directory"}, 400)
                key = os.path.normcase(os.path.abspath(p))
                arr = [x for x in read_custom_roots()
                       if os.path.normcase(os.path.abspath(x)) != key]
                arr.append(os.path.normpath(p))
                write_custom_roots(arr)
                # 回读：返回的是截断后的真实列表，前端据此把被挤掉的旧目录从下拉里移除
                return self._json({"roots": read_custom_roots(), "max": MAX_CUSTOM_ROOTS})

            if u.path == "/api/browse-folder":
                # 模态对话框只阻塞当前请求线程；服务是 ThreadingHTTPServer，
                # 会话保存、托盘轮询等其它请求不受影响。
                p = browse_folder()
                if not p or not os.path.isdir(p):
                    return self._json({"cancelled": True})
                return self._json({"path": os.path.normpath(p)})

            if u.path == "/api/classify":
                return self._json({"kind": kind_of(body.get("path", ""))})

            if u.path == "/api/save":
                p = body.get("path", "")
                text = body.get("text")
                enc = body.get("encoding") or "utf-8"
                if not os.path.isfile(p) or text is None:
                    return self._json({"error": "bad request"}, 400)
                # utf-8-sig / utf-16 自带 BOM，不要重复添加
                data = text.encode(enc, errors="replace")
                bak = p + ".uniview-bak"
                if not os.path.exists(bak):
                    try:
                        with open(p, "rb") as f, open(bak, "wb") as g:
                            g.write(f.read())
                    except OSError:
                        pass
                tmp = p + ".uniview-tmp"
                with open(tmp, "wb") as f:
                    f.write(data)
                os.replace(tmp, p)
                return self._json({"ok": True, "bytes": len(data), "backup": bak,
                                   "encoding": enc})

            if u.path == "/api/office/embed":
                return self._json(HOST.embed(body.get("path", ""),
                                             body.get("rect", {"x": 0, "y": 0, "w": 800, "h": 600}),
                                             body.get("scale", 1.0)))

            if u.path == "/api/office/sync":
                HOST.sync(body.get("rect", {}), body.get("scale", 1.0))
                return self._json(HOST.info())

            if u.path == "/api/office/detach":
                return self._json(HOST.detach())

            if u.path == "/api/open-external":
                p = body.get("path", "")
                if not os.path.isfile(p):
                    return self._json({"error": "not a file"}, 404)
                HOST.detach()
                os.startfile(p)
                return self._json({"ok": True, "message": "已交给系统默认程序打开"})

            if u.path == "/api/quit":
                # 真正的退出：停服务、摘托盘图标、清运行标记，整个进程结束。
                # session.json 保留不动，下次启动照样恢复上次的标签。
                threading.Thread(target=shutdown_app, daemon=True).start()
                return self._json({"ok": True})

            return self._json({"error": "not found"}, 404)
        except Exception as ex:
            traceback.print_exc()
            return self._json({"error": str(ex)}, 500)


# ---------------------------------------------------------------- 启动

def find_browser():
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def write_runtime(port, token):
    try:
        with open(RUNTIME_PATH, "w", encoding="utf-8") as f:
            json.dump({"port": port, "token": token, "pid": os.getpid()}, f)
    except OSError:
        pass


def read_runtime():
    try:
        with open(RUNTIME_PATH, encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d.get("port"), int) and d.get("token"):
            return d
    except Exception:
        pass
    return None


# ---------------------------------------------------------------- 会话保存/恢复

def read_session():
    """读出上次保存的标签列表；已移动/删除的文件在这里就过滤掉。"""
    try:
        with open(SESSION_PATH, encoding="utf-8") as f:
            s = json.load(f)
    except Exception:
        return None
    if not isinstance(s, dict) or not isinstance(s.get("tabs"), list):
        return None
    tabs = [p for p in s["tabs"] if isinstance(p, str) and os.path.isfile(p)]
    if not tabs:
        return None
    if not os.path.isfile(str(s.get("active", ""))):
        s["active"] = tabs[-1]
    sp = s.get("split")
    if isinstance(sp, dict):
        if not os.path.isfile(str(sp.get("left", ""))) or not os.path.isfile(str(sp.get("right", ""))):
            sp["on"] = False
    s["tabs"] = tabs
    return s


def write_session(data):
    """前端周期性上报当前标签状态；只收白名单字段，原子写避免半截文件。"""
    if not isinstance(data, dict):
        return
    tabs = [p for p in data.get("tabs", []) if isinstance(p, str)][:64]
    sp = data.get("split")
    split = {"on": False, "left": "", "right": "", "focus": "L", "pct": 50}
    if isinstance(sp, dict):
        split.update({
            "on": bool(sp.get("on")),
            "left": str(sp.get("left", "")),
            "right": str(sp.get("right", "")),
            "focus": "L" if sp.get("focus") == "L" else "R",
            "pct": max(20, min(80, int(sp.get("pct", 50) or 50))),
        })
    out = {"v": 1, "tabs": tabs, "active": str(data.get("active", "")),
           "split": split, "savedAt": data.get("savedAt")}
    tmp = SESSION_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False)
        os.replace(tmp, SESSION_PATH)
    except OSError as ex:
        log("SESSION save failed | " + str(ex))



def ping_instance(info):
    """已有服务还活着吗？避免每次「打开方式」都再起一个服务进程。"""
    try:
        u = "http://127.0.0.1:%d/api/ping?t=%s" % (
            int(info["port"]), urllib.parse.quote(info["token"], safe=""))
        with urllib.request.urlopen(u, timeout=1.5) as f:
            return json.loads(f.read().decode("utf-8")).get("ok") is True
    except Exception:
        return False


def open_window(url, no_window=False):
    exe = None if no_window else find_browser()
    if no_window:
        print("仅服务模式，手动访问：" + url)
        return
    if exe:
        profile = os.path.join(os.environ.get("TEMP", ROOT), "uniview-profile")
        args = [exe, "--app=" + url, "--window-size=1440,920",
                "--window-position=60,40", "--user-data-dir=" + profile,
                "--no-first-run", "--no-default-browser-check",
                "--disable-features=msEdgeSidebarV2"]
        try:
            subprocess.Popen(args, close_fds=True)
            return
        except Exception as ex:
            print("独立窗口启动失败，改用默认浏览器:", ex)
    webbrowser.open(url)


def ensure_window_open():
    """窗口活着就拉到最前；已关闭（或长期没轮询）就重新开一个。
    返回 True 表示新开了窗口。"""
    if STATE["seen"] <= 0 or time.time() - STATE["seen"] > 90:
        STATE["seen"] = time.time()   # 先标记，避免并发重复开窗
        url = "http://127.0.0.1:%d/?t=%s" % (PORT, TOKEN)
        threading.Thread(target=open_window, args=(url,), daemon=True).start()
        log("WINDOW reopen")
        return True
    raise_window()
    return False


def shutdown_app():
    """真正退出整个软件：先给前端 0.4 秒收到响应，再收尾并结束进程。"""
    time.sleep(0.4)
    log("QUIT | 用户主动退出")
    try:
        HOST.detach()
    except Exception:
        pass
    try:
        TRAY.remove()
    except Exception:
        pass
    try:
        HTTPD.shutdown()
    except Exception:
        pass
    try:
        os.remove(RUNTIME_PATH)
    except OSError:
        pass
    os._exit(0)


def queue_to_instance(info, path):
    """把「打开方式」的文件投给已运行的窗口，由它开成新标签，不另开窗口。"""
    try:
        url = "http://127.0.0.1:%d/api/queue?t=%s" % (
            int(info["port"]), urllib.parse.quote(info["token"], safe=""))
        req = urllib.request.Request(
            url,
            data=json.dumps({"path": os.path.abspath(path)}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST")
        with urllib.request.urlopen(req, timeout=2.5) as f:
            return json.loads(f.read().decode("utf-8")).get("ok") is True
    except Exception as ex:
        log("QUEUE failed | " + str(ex))
        return False


def reopen_instance(info):
    """已在托盘驻留时再次双击图标：不重起服务，让现有进程把窗口开回来。"""
    try:
        url = "http://127.0.0.1:%d/api/reopen?t=%s" % (
            int(info["port"]), urllib.parse.quote(info["token"], safe=""))
        req = urllib.request.Request(url, data=b"{}",
                                     headers={"Content-Type": "application/json"},
                                     method="POST")
        with urllib.request.urlopen(req, timeout=2.5) as f:
            return json.loads(f.read().decode("utf-8")).get("ok") is True
    except Exception as ex:
        log("REOPEN failed | " + str(ex))
        return False


def launch(no_window=False, file_path=None):
    global HTTPD, PORT
    # 已有实例活着（包括窗口已关、托盘驻留中）：不再起第二个服务
    for attempt in range(4):
        info = read_runtime()
        if info and ping_instance(info):
            if file_path:
                log("OPENWITH reuse | " + file_path)
                queue_to_instance(info, file_path)
            else:
                log("REOPEN existing instance")
                reopen_instance(info)
            return
        if attempt < 3:
            time.sleep(0.2)   # 第一个进程可能刚写完 .runtime.json，稍等再试

    HTTPD = Server(("127.0.0.1", 0), Handler)
    PORT = HTTPD.server_address[1]
    write_runtime(PORT, TOKEN)
    # 供安装程序识别的运行互斥体（安装/卸载时 Inno Setup 据此提示先关闭软件）
    if sys.platform == "win32":
        try:
            k32.CreateMutexW(None, False, "UniViewAppMutex")
        except Exception:
            pass
    # 窗口马上就要打开：先标记为「活着」，避免并发的第二个文件误判成无窗口而多开一个
    STATE["seen"] = time.time()
    url = "http://127.0.0.1:%d/?t=%s" % (PORT, TOKEN)
    if file_path:
        url += "&file=" + urllib.parse.quote(file_path, safe="")
        log("OPENWITH new instance | " + file_path)
    print("UniView 已启动: %s" % url)
    print("关闭窗口只是收起到托盘；托盘图标右键可真正退出。")
    threading.Thread(target=HTTPD.serve_forever, daemon=True).start()

    HOST.start_follower()
    threading.Thread(target=lambda: HOST.wait_host(25), daemon=True).start()
    TRAY.start()
    open_window(url, no_window)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        HOST.detach()
        TRAY.remove()
        HTTPD.shutdown()
        try:
            os.remove(RUNTIME_PATH)
        except OSError:
            pass
        print("\n已退出。")


# ---------------------------------------------------------------- 无 GUI 自测

def self_test():
    samples = os.path.normpath(os.path.join(ROOT, "..", "web-viewer", "samples"))
    ok = []
    if os.path.isdir(samples):
        files = walk_files(samples)
        ok.append(("目录扫描", len(files) > 0, "%d 个可查看文件" % len(files)))
        for n in ("GBK文本.txt", "纯文本.txt", "示例说明.md"):
            p = os.path.join(samples, n)
            if os.path.isfile(p):
                t, e = read_text(p)
                ok.append((n + " 解码", len(t) > 0 and "锟" not in t, "编码=%s 字数=%d" % (e, len(t))))
        for n in ("示例文档.docx", "示例表格.xlsx", "示例演示.pptx"):
            p = os.path.join(samples, n)
            if os.path.isfile(p):
                t = extract_text(p)
                ok.append((n + " 正文抽取", len(t) > 0, "%d 字" % len(t)))
        hits = []
        for p in walk_files(samples):
            t = extract_text(p)
            if "示例" in t:
                hits.append(os.path.basename(p))
        ok.append(("跨格式搜索命中", len(hits) >= 2, ", ".join(hits[:6])))
    ok.append(("WPS 组件定位", find_wps_exe(".docx") is not None,
               str(find_wps_exe(".docx"))))
    ok.append(("浏览器定位", find_browser() is not None, str(find_browser())))
    ok.append(("格式分类", kind_of("a.md") == "text" and kind_of("a.docx") == "office"
               and kind_of("a.pdf") == "pdf" and kind_of("a.png") == "image", "text/office/pdf/image"))
    bad = 0
    for name, passed, detail in ok:
        print(("  PASS  " if passed else "  FAIL  ") + name + "  |  " + detail)
        bad += 0 if passed else 1
    print("\n%d 项通过 / %d 项失败" % (len(ok) - bad, bad))
    return 1 if bad else 0


def smoke():
    """起一个真实服务，用 HTTP 走一遍全部 API（不弹窗口，可在本会话验证）。"""
    global HTTPD, PORT
    import urllib.error
    HTTPD = Server(("127.0.0.1", 0), Handler)
    PORT = HTTPD.server_address[1]
    threading.Thread(target=HTTPD.serve_forever, daemon=True).start()
    base = "http://127.0.0.1:%d" % PORT

    def call(path, body=None, raw=False):
        data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
        req = urllib.request.Request(base + path, data=data,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as f:
            b = f.read()
            return b if raw else json.loads(b.decode("utf-8"))

    samples = os.path.normpath(os.path.join(ROOT, "..", "web-viewer", "samples"))
    t = "?t=" + TOKEN
    ok = []

    def check(name, fn):
        try:
            detail = fn()
            ok.append((name, True, str(detail)))
        except Exception as ex:
            ok.append((name, False, repr(ex)))
        return ok[-1]

    check("首页 HTML", lambda: len(call("/", raw=True)))
    check("离线解析库", lambda: len(call("/vendor/marked.min.js", raw=True)))
    check("离线图表库", lambda: len(call("/vendor/mermaid.min.js", raw=True)))
    check("启动配置 /api/boot", lambda: len(call("/api/boot" + t)["roots"]))
    check("启动配置含记住的目录", lambda: "customRoots=%s" % (call("/api/boot" + t).get("customRoots", "MISSING"),))
    def roots_limit():
        """造 20 个真实目录，验证只保留最近 15 个（用临时文件，不碰用户数据）"""
        import shutil
        base = os.path.join(tempfile.gettempdir(), "uniview_roots_limit")
        shutil.rmtree(base, ignore_errors=True)
        dirs = []
        for i in range(20):
            d = os.path.join(base, "d%02d" % i)
            os.makedirs(d, exist_ok=True)
            dirs.append(d)
        global ROOTS_PATH
        old = ROOTS_PATH
        try:
            ROOTS_PATH = os.path.join(base, "roots.json")
            write_custom_roots(dirs)
            got = read_custom_roots()
            assert len(got) == MAX_CUSTOM_ROOTS, got
            assert got == dirs[-MAX_CUSTOM_ROOTS:], got
            return "20 → %d，保留最后 %d 个" % (len(got), MAX_CUSTOM_ROOTS)
        finally:
            ROOTS_PATH = old
            shutil.rmtree(base, ignore_errors=True)
    check("记住目录上限 15", roots_limit)
    try:
        call("/api/boot")
        ok.append(("未授权请求被拒", False, "401 未触发"))
    except urllib.error.HTTPError as ex:
        ok.append(("未授权请求被拒", ex.code == 401, "HTTP %d" % ex.code))
    check("目录树 /api/tree", lambda: len(call("/api/tree?path=" + urllib.parse.quote(samples) + "&t=" + TOKEN)["items"]))
    check("此电脑磁盘列表", lambda: "%d 个盘符" % len(call("/api/tree?path=" + urllib.parse.quote(THIS_PC) + "&t=" + TOKEN)["items"]))
    check("目录存在性探测", lambda: "exists=%s" % call("/api/tree?path=" + urllib.parse.quote(os.path.expanduser("~")) + "&probe=1&t=" + TOKEN)["exists"])
    def tree_filter():
        d = os.path.join(tempfile.gettempdir(), "uniview_odd_dir")
        os.makedirs(d, exist_ok=True)
        odd = os.path.join(d, "uniview_odd.xyz")
        txt = os.path.join(d, "uniview_ok.txt")
        for p in (odd, txt):
            with open(p, "w", encoding="utf-8") as f:
                f.write("x")
        try:
            base = "/api/tree?path=" + urllib.parse.quote(d) + "&t=" + TOKEN
            r = call(base)
            names = [i["name"] for i in r["items"]]
            assert "uniview_odd.xyz" not in names, "默认仍在列出无法打开的文件"
            assert "uniview_ok.txt" in names, "可打开的文本被误隐藏"
            assert r["hidden"] == 1, "hidden 计数不对"
            a = call(base + "&all=1")
            assert any(i["name"] == "uniview_odd.xyz" and i["viewable"] is False
                       for i in a["items"]), "all=1 未列出全部文件"
            return "默认隐藏 %d 个不可打开文件；all=1 可全列出" % r["hidden"]
        finally:
            os.remove(odd)
            os.remove(txt)
            os.rmdir(d)
    check("目录树默认隐藏无法打开的文件", tree_filter)
    check("类型判定 /api/classify", lambda: call("/api/classify" + t, {"path": "a.docx"})["kind"])
    check("GBK 文本读取", lambda: call("/api/text?path=" + urllib.parse.quote(os.path.join(samples, "GBK文本.txt")) + "&t=" + TOKEN)["encoding"] + " / 无乱码=" + str("锟" not in call("/api/text?path=" + urllib.parse.quote(os.path.join(samples, "GBK文本.txt")) + "&t=" + TOKEN)["text"]))
    check("跨格式搜索", lambda: len(call("/api/search?q=" + urllib.parse.quote("示例") + "&root=" + urllib.parse.quote(samples) + "&t=" + TOKEN)["hits"]))
    check("PDF 原字节", lambda: call("/api/raw?path=" + urllib.parse.quote(os.path.join(samples, "示例.pdf")) + "&t=" + TOKEN, raw=True)[:5])

    def favicon_check():
        with urllib.request.urlopen(base + "/favicon.ico", timeout=5) as r:
            b = r.read()
            ctype = r.headers.get("Content-Type", "")
        assert len(b) > 1000 and ctype == "image/x-icon", "图标缺失或类型不对"
        return "%d 字节 / %s" % (len(b), ctype)
    check("窗口图标 /favicon.ico", favicon_check)

    def html_sandbox():
        d = os.path.join(tempfile.gettempdir(), "uniview_smoke_html")
        os.makedirs(d, exist_ok=True)
        hp = os.path.join(d, "page.html")
        cp = os.path.join(d, "a.css")
        with open(hp, "w", encoding="utf-8") as f:
            f.write('<!DOCTYPE html><html><head><link rel="stylesheet" href="a.css"></head>'
                    '<body><h1>OK</h1></body></html>')
        with open(cp, "w", encoding="utf-8") as f:
            f.write("h1{color:red}")
        try:
            key = call("/api/boot" + t)["htmlKey"]
            assert key and key != TOKEN, "HTML 钥匙应与 API 令牌分开"
            enc = base64.urlsafe_b64encode(hp.encode("utf-8")).decode("ascii").rstrip("=")
            u = "/html/%s/%s" % (key, enc)
            with urllib.request.urlopen(base + u, timeout=5) as r:
                body = r.read().decode("utf-8")
                csp = r.headers.get("Content-Security-Policy", "")
            assert "<base href=" in body and "<meta charset=\"utf-8\">" in body, "base/charset 注入缺失"
            assert "connect-src 'none'" in csp and "form-action 'none'" in csp, "CSP 未禁止联网"
            assert b"color:red" in urllib.request.urlopen(base + u + "/a.css", timeout=5).read(), "同目录 CSS 代理失败"
            assert call("/api/classify" + t, {"path": hp})["kind"] == "html", "HTML 类型判定错误"
            try:
                urllib.request.urlopen(base + "/html/WRONGKEY/" + enc, timeout=5)
                raise AssertionError("错误钥匙未被拒绝")
            except urllib.error.HTTPError as ex:
                assert ex.code == 403, "错误钥匙应返回 403"
            return "CSP 禁止联网；base + 本地 CSS 代理正常；错误钥匙被拒"
        finally:
            for p in (hp, cp):
                try:
                    os.remove(p)
                except OSError:
                    pass
            try:
                os.rmdir(d)
            except OSError:
                pass
    check("HTML 隔离渲染与资源代理", html_sandbox)
    check("嵌入状态查询", lambda: call("/api/office" + t)["status"])

    tmp = os.path.join(tempfile.gettempdir(), "uniview_smoke.txt")
    with open(tmp, "wb") as f:
        f.write("原始内容".encode("gb18030"))
    def save_roundtrip():
        d = call("/api/save" + t, {"path": tmp, "text": "改过的中文", "encoding": "gb18030"})
        back = open(tmp, "rb").read().decode("gb18030")
        os.remove(tmp)
        if os.path.exists(tmp + ".uniview-bak"):
            os.remove(tmp + ".uniview-bak")
        return "bytes=%s 回读=%s 备份=%s" % (d["bytes"], back, d["ok"])
    check("保存写回原编码 + 备份", save_roundtrip)

    HTTPD.shutdown()
    bad = 0
    for name, passed, detail in ok:
        print(("  PASS  " if passed else "  FAIL  ") + name + "  |  " + detail)
        bad += 0 if passed else 1
    print("\n%d 项通过 / %d 项失败" % (len(ok) - bad, bad))
    return 1 if bad else 0


if __name__ == "__main__":
    # pythonw.exe 无控制台，sys.stdout/stderr 为 None，print() 会崩；统一重定向到日志
    if sys.stdout is None or sys.stderr is None:
        try:
            sink = open(LOG_PATH, "a", encoding="utf-8", errors="replace")
        except OSError:
            sink = open(os.devnull, "w", encoding="utf-8")
        if sys.stdout is None:
            sys.stdout = sink
        if sys.stderr is None:
            sys.stderr = sink
    ap = argparse.ArgumentParser(description="UniView 本地文件统一查看器")
    ap.add_argument("--self-test", action="store_true", help="只做无界面逻辑自测")
    ap.add_argument("--smoke", action="store_true", help="起真实服务走一遍 HTTP API")
    ap.add_argument("--no-window", action="store_true", help="只起服务，不开窗口")
    ap.add_argument("--file", default="", help="启动时直接打开该文件（供系统「打开方式」调用）")
    a = ap.parse_args()
    if a.self_test:
        sys.exit(self_test())
    if a.smoke:
        sys.exit(smoke())
    launch(a.no_window, os.path.abspath(a.file) if a.file else None)
