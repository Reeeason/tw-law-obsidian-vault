"""分科六法庫：把全國法規資料庫下載的 Word 檔（或由它轉出的乾淨 md 檔）轉成 Obsidian 逐條頁面。

子命令（一律在 tools/ 資料夾內用 python -I 執行；--src、--dry-run 的相對路徑以目前工作目錄為準）：
  scan       掃描法規，產生設定檔草稿，列出需要人工補的項目
  build      產生或更新逐條頁面、法規總覽，並自動核對
  check-ext  把設定檔裡的外部引用條文和全國法規資料庫官網逐段比對
  verify     只做核對，不寫檔

用法：
  python -I 法規庫.py scan      --src 法規.docx（或 法規.md）
  python -I 法規庫.py check-ext --law 法規全名
  python -I 法規庫.py build     --src 法規.docx --law 法規全名 [--dry-run 暫存資料夾]
  python -I 法規庫.py verify    --src 法規.docx --law 法規全名

設定與登錄：
  法規登錄.json             每部法規一筆：簡稱、分類資料夾、設定檔路徑（選填 src：來源檔，相對於本資料夾）
  設定/<法規全名>.json       自擬條名、外部引用條文、效力異動
頁面規則：
  條文區在「%% 條文區開始 ... %%」到「%% 條文區結束 %%」之間，由本腳本整段覆寫；
  條文區以外是使用者筆記，更新時絕不改動。屬性區只更新系統欄位，保留使用者欄位。
"""
import argparse, datetime, json, os, re, sys, urllib.request, html

HERE = os.path.dirname(os.path.abspath(__file__))
# 頁面輸出位置：預設為本庫的「分科六法整理」資料夾；要寫到別處時設定環境變數 LAW_LIBRARY_DIR
VAULT = os.environ.get("LAW_LIBRARY_DIR") or os.path.join(os.path.dirname(HERE), "分科六法整理")
REGISTRY = os.path.join(HERE, "法規登錄.json")
DASHBOARD = os.path.join(VAULT, "00-總覽與看板", "分科六法總覽.md")
CATEGORIES = ["01-公法", "02-民事法", "03-刑事法", "04-商事法與智慧財產權法", "05-憲法法庭判決及大法官解釋"]
START = "%% 條文區開始：本區由腳本產生，修法更新時會整段覆寫，筆記請寫在條文區以外 %%"
END = "%% 條文區結束 %%"
LEVEL = {"編": 2, "章": 3, "節": 4, "款": 5, "目": 6}
LEVEL_KEY = {"編": "part", "章": "chapter", "節": "section", "款": "subsection", "目": "item"}
SYSTEM_KEYS = ["type", "law", "law_short", "article", "seq", "caption", "caption_draft",
               "part", "chapter", "section", "subsection", "item", "status", "status_note", "source"]
USER_DEFAULTS = {"progress": "未整理", "exams": []}
BOUNDARY = set("、，。；：「」（）(或及與依犯之準用照按對於為以同並且酌就過反受有因")


def die(msg):
    sys.exit("錯誤：" + msg)


def load_json(p, default=None):
    if not os.path.exists(p):
        if default is not None: return default
        die(f"找不到 {p}")
    return json.load(open(p, encoding="utf-8"))


def save_json(p, obj):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, "w", encoding="utf-8", newline="\n").write(json.dumps(obj, ensure_ascii=False, indent=1) + "\n")


# ---------------- 中文數字 ----------------
DIGIT = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def cn(s):
    n, cur = 0, 0
    for ch in s:
        if ch in DIGIT: cur = DIGIT[ch]
        elif ch == "十": n += (cur or 1) * 10; cur = 0
        elif ch == "百": n += (cur or 1) * 100; cur = 0
        elif ch == "千": n += (cur or 1) * 1000; cur = 0
    return n + cur


# ---------------- 解析 Word 檔 ----------------
HEAD = re.compile(r"^第 (\S+) ([編章節款目])(之[一二三四五六七八九十]+)?\s+(.+)$")
ART = re.compile(r"^第 ([0-9]+(?:-[0-9]+)?) 條[ \t]*(.*)$")


def squash(t):  # 「第 十一 章 公共危險罪」→「第十一章 公共危險罪」
    m = HEAD.match(t)
    return f"第{m.group(1)}{m.group(2)}{m.group(3) or ''} {m.group(4).strip()}"


def read_source(path):
    """Word 檔或 md／txt 檔都轉成段落清單；md 的「#」標題記號會拿掉。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        import docx
        return [p.text.rstrip() for p in docx.Document(path).paragraphs]
    if ext in (".md", ".txt"):
        t = open(path, encoding="utf-8-sig").read()
        if "> 備註" in t or "[[" in t or re.search(r"條 \([^)]*\)", re.sub(r"^#.*$", "", t, flags=re.M)):
            die("這份 md 看起來是引用標註版或已轉成 Obsidian 的版本，請改用未標註的乾淨版（docx 直接轉出的那份）")
        return [re.sub(r"^#+ ", "", l.rstrip()) for l in t.split("\n")]
    die(f"不支援的檔案類型 {ext}，請給 .docx 或 .md")


def parse_docx(path):
    raw = read_source(path)
    paras = []
    for t in raw:
        if not t.strip(): continue
        m = re.match(r"^( +)(.*)$", t)
        if m: t = "　" * len(m.group(1)) + m.group(2)  # 款次縮排改全形空格
        paras.append(t)
    law = {"title": None, "download": "", "revised": "", "revised_label": "修正日期", "history": [], "arts": {}, "order": [], "toc": []}
    i = 0
    while i < len(paras) and not (HEAD.match(paras[i]) or ART.match(paras[i])):
        t = paras[i]
        if t.startswith("下載日期"): law["download"] = t.split("：", 1)[-1].strip()
        elif re.match(r"^(修正|修訂|公布|制定)日期", t):
            law["revised_label"] = t[:4]; law["revised"] = re.sub(r"^\S\S日期[：:\s]*", "", t).strip()
        elif re.match(r"^\d+\.", t): law["history"].append(t)
        elif law["title"] is None and not t.endswith("：") and "日期" not in t: law["title"] = t.strip()
        i += 1
    path_now = {}
    cur = None
    for t in paras[i:]:
        h = HEAD.match(t)
        if h:
            lv = h.group(2)
            for k in list(path_now):
                if LEVEL[k] >= LEVEL[lv]: del path_now[k]
            path_now[lv] = squash(t)
            law["toc"].append(("h", LEVEL[lv], path_now[lv])); continue
        a = ART.match(t)
        if a:
            cur = a.group(1)
            cap = re.sub(r"^[（(]\s*|\s*[)）]$", "", a.group(2).strip()).strip()
            law["arts"][cur] = {"caption": cap, "path": dict(path_now), "paras": [], "note": []}
            law["order"].append(cur); law["toc"].append(("a", cur)); continue
        if cur is None: die(f"條文前出現無法辨識的段落：{t[:40]}")
        art = law["arts"][cur]
        if re.match(r"^編\s*註：", t) or art["note"]:
            art["note"].append(t); continue
        art["paras"].append(t)
    for a in law["arts"].values():
        a["deleted"] = a["caption"] == "刪除" or a["paras"] in (["（刪除）"], ["(刪除)"])  # 保險法第168-7條用半形括號
        if a["deleted"]: a["caption"] = ""
    if not law["order"]: die("沒有解析到任何條文，請確認是全國法規資料庫下載的 Word 檔，或由它轉出的 md 檔")
    return law


# ---------------- 登錄與既有頁面 ----------------
def registry():
    return load_json(REGISTRY, {})


def law_dir(full, reg):
    r = reg[full]
    return os.path.join(VAULT, r["category"], full)


def page_name(full, no): return f"{full}第{no}條"


_SRC_CACHE = {}


def source_caption(full, no, reg):
    """已登錄、尚未建頁但登錄表有 src 的法規：從來源檔取條名（含自擬）；查無該條回傳 None。"""
    r = reg[full]
    if not r.get("src"): return None
    if full not in _SRC_CACHE:
        _SRC_CACHE[full] = (parse_docx(os.path.join(HERE, r["src"])), load_json(os.path.join(HERE, r["config"]), {}).get("labels", {}))
    law, labels = _SRC_CACHE[full]
    a = law["arts"].get(no)
    if a is None: return None
    if a["deleted"]: return "已刪除"
    if a["caption"]: return a["caption"]
    return "〔自擬〕" + labels[no] if labels.get(no) else ""


def existing_caption(full, no, reg):
    """其他法規已建頁面的條名（含自擬）；沒有頁面時改從登錄的來源檔取；都沒有回傳 None。"""
    if full not in reg: return None
    p = os.path.join(law_dir(full, reg), page_name(full, no) + ".md")
    if not os.path.exists(p): return source_caption(full, no, reg)
    fm = read_frontmatter(open(p, encoding="utf-8").read())[0]
    if fm.get("status") == "已刪除": return "已刪除"
    if fm.get("caption"): return fm["caption"]
    if fm.get("caption_draft"): return "〔自擬〕" + fm["caption_draft"]
    return ""


def read_frontmatter(text):
    import yaml
    if not text.startswith("---\n"): return {}, text
    end = text.index("\n---\n", 4)
    return yaml.safe_load(text[4:end]) or {}, text[end + 5:]


def dump_frontmatter(d):
    out = ["---"]
    for k, v in d.items():
        if isinstance(v, list):
            if not v: out.append(f"{k}: []")
            else:
                out.append(f"{k}:")
                out += [f"  - {json.dumps(x, ensure_ascii=False)}" for x in v]
        elif isinstance(v, int) and not isinstance(v, bool):
            out.append(f"{k}: {v}")
        else:
            out.append(f"{k}: {json.dumps('' if v is None else str(v), ensure_ascii=False)}")
    out.append("---")
    return "\n".join(out)


# ---------------- 引用解析 ----------------
NUM = "零一二三四五六七八九十百千"
CIT = re.compile(rf"第([{NUM}]+)條(?:之([一二三四五六七八九十]+))?"
                 rf"(?:第[{NUM}]+項)?(?:第[{NUM}]+款)?(?:前段|後段|但書)?"
                 rf"(?:[、及或至]?第[{NUM}]+[項款](?:前段|後段|但書)?)*")  # 「第一項至第三項」「第二項第一款、第三款」整串算同一條
GENERIC_LAW = re.compile(r"(?:法|條例|通則|規則|規程|辦法|細則|公約|協定|條文)$")  # 「條文」：如「憲法增修條文」
PREV = re.compile(r"前([二三四五六七八九])?條")
CNUM = "零一二三四五六七八九"
JOIN = ("、", "及", "或", "至", "與")
UNKNOWN = "?"  # render() 中「前一個條號的法規判不出來」的記號


class Ctx:
    def __init__(self, law, full, short, cfg, reg):
        self.law, self.full, self.short, self.cfg, self.reg = law, full, short, cfg, reg
        self.arts, self.order = law["arts"], law["order"]
        self.labels = cfg.get("labels", {})
        self.ext = cfg.get("external", {})
        names = {}
        for f, r in reg.items():
            names[f] = f
            if r.get("short"): names[r["short"]] = f
        for k in self.ext: names[k.split("|")[0]] = k.split("|")[0]
        for k, v in cfg.get("law_names", {}).items(): names[k] = v
        self.names = sorted(names.items(), key=lambda x: -len(x[0]))
        self.problems = {"unknown_law": [], "missing_ext": [], "missing_label": set(), "ext_range": []}
        self.stats = {"cite": 0, "prev": 0, "scope": 0, "range": 0, "ext_note": 0, "ext_link": 0, "ext_range": 0}

    def label(self, no, strict=True):
        a = self.arts[no]
        if a["deleted"]: return "已刪除"
        if a["caption"]: return a["caption"]
        if no in self.labels and self.labels[no]: return "〔自擬〕" + self.labels[no]
        if strict: self.problems["missing_label"].add(no)
        return ""

    def link(self, no, shown, full=None):
        return f"[[{page_name(full or self.full, no)}|{shown}]]"

    def which_law(self, prefix, gap, last_law):
        """回傳 (法規名或 None 表本法, 是否無法判定)"""
        if prefix.endswith("本法"): return None, False
        if prefix.endswith("同法"): return (None, True) if last_law in (None, UNKNOWN) else (last_law, False)
        for name, full in self.names:
            if prefix.endswith(name):
                before = prefix[:-len(name)][-1:] if len(prefix) > len(name) else ""
                # 只有「刑法」「民法」這類兩字名稱，才可能是更長法規名的尾巴（「陸海空軍刑法」）
                if len(name) <= 2 and before and before not in BOUNDARY and not before.isascii():
                    return None, True
                return (None if full == self.full else full), False
        if GENERIC_LAW.search(prefix): return None, True
        # 列舉中的後續條號（「刑法第221條之強制性交罪、第222條之…」）沿用前一部外部法規；
        # 前一部判不出來（UNKNOWN）時，後續條號也判不出來，不可退回本法
        if last_law and (gap in JOIN or re.search(r"[、，及或與]$", gap) and not re.search(r"本法|[。；：]", gap)):
            return (None, True) if last_law == UNKNOWN else (last_law, False)
        return None, False

    def ext_note(self, law, no, cur, notes):
        """尚未建頁的外部法規條文：附收合備註框。"""
        e = self.ext.get(f"{law}|{no}")
        if e is None:
            self.problems["missing_ext"].append((cur, f"{law}|{no}")); return
        if e.get("omit"):  # 引用的是舊版條號，現行同條號內容不同：不附條文，只附說明
            notes.append(("note-", f"{law}第{no}條（未附現行條文）", [e["note"]]))
        else:
            notes.append(("note-", f"{law}第{no}條 (〔自擬〕{e['label']})" if e.get("draft", True) else f"{law}第{no}條 ({e['label']})",
                          e["paras"] + ([f"〔說明〕{e['note']}"] if e.get("note") else [])))
        self.stats["ext_note"] += 1

    def render(self, l, cur):
        edits, notes = [], []
        last_end, last_law, last_no, last_txt = -10, None, None, None
        last_ext, last_txt_ext = None, None
        for mm in CIT.finditer(l):
            no = str(cn(mm.group(1))) + (f"-{cn(mm.group(2))}" if mm.group(2) else "")
            txt_no = "第" + mm.group(1) + "條" + ("之" + mm.group(2) if mm.group(2) else "")
            gap = l[last_end:mm.start()]
            law, unknown = self.which_law(l[:mm.start()], gap, last_law)
            if unknown:
                self.problems["unknown_law"].append((cur, l[max(0, mm.start() - 20):mm.end()]))
                last_end, last_law, last_no = mm.end(), UNKNOWN, None
                continue
            if gap == "至" and (law or last_law):
                key = f"{law}|{last_ext[1]}-{no}" if law and last_ext and last_ext[0] == law else None
                rng = self.cfg.get("ext_ranges", {}).get(key) if key else None
                if rng is None:
                    self.problems["ext_range"].append((cur, key or l[max(0, mm.start() - 30):mm.end()]))
                elif isinstance(rng, list):  # 範圍內每一條（兩端另外處理）
                    mids = [t for t in rng if t not in (last_ext[1], no)]
                    if existing_caption(law, no, self.reg) is not None:
                        notes.append(("note", f"範圍：{law}{last_txt_ext}至{txt_no}",
                                      ["、".join(f"{self.link(t, '第' + t + '條', law)} {existing_caption(law, t, self.reg) or ''}".rstrip() for t in rng)]))
                    else:
                        for t in mids: self.ext_note(law, t, cur, notes)
                    self.stats["ext_range"] += 1
                # rng 為 {"omit": true, "note": …}：範圍引用的是舊版條號，兩端的說明已足，不另附
            elif not law and gap == "至" and last_no and no not in self.arts:
                self.problems["unknown_law"].append((cur, f"本法範圍的終點不存在：{last_txt}至{txt_no}"))
            elif not law and gap == "至" and last_no:
                span = self.order[self.order.index(last_no):self.order.index(no) + 1]
                notes.append(("note", f"範圍：{last_txt}至{txt_no}",
                              ["、".join(f"{self.link(t, '第' + t + '條')} {self.label(t)}".rstrip() for t in span)]))
                self.stats["range"] += 1
            last_end, last_law = mm.end(), law
            last_no, last_txt = (None, None) if law else (no, txt_no)
            last_ext, last_txt_ext = ((law, no), txt_no) if law else (None, None)
            if law:
                cap = existing_caption(law, no, self.reg)
                if cap is not None:  # 該法規已建頁面：直接連過去
                    edits.append((mm.start(), mm.end(), f"{self.link(no, mm.group(0), law)}" + (f" ({cap})" if cap else "")))
                    self.stats["ext_link"] += 1
                else:
                    self.ext_note(law, no, cur, notes)
                continue
            if no == cur: continue
            if no not in self.arts:
                self.problems["unknown_law"].append((cur, f"本法查無第{no}條：" + mm.group(0))); continue
            lab = self.label(no)
            edits.append((mm.start(), mm.end(), self.link(no, mm.group(0)) + (f" ({lab})" if lab else "")))
            self.stats["cite"] += 1
        for mm in PREV.finditer(l):
            k = CNUM.index(mm.group(1)) if mm.group(1) else 1
            idx = self.order.index(cur)
            tg = self.order[max(0, idx - k):idx]
            shown = self.link(tg[0], mm.group(0)) if k == 1 else mm.group(0)
            parts = [f"{self.link(t, '第' + t + '條')} {self.label(t)}".rstrip() for t in tg]
            edits.append((mm.start(), mm.end(), f"{shown} (" + "；".join(parts) + ")")); self.stats["prev"] += 1
        for unit in ("編", "章", "節"):
            mm = re.search("本" + unit, l)
            if mm and unit in self.arts[cur]["path"]:
                h = self.arts[cur]["path"][unit]
                edits.append((mm.start(), mm.end(), f"本{unit} ([[{self.full}#{h}|{h}]])")); self.stats["scope"] += 1
        for s, e, r in sorted(edits, reverse=True):
            l = l[:s] + r + l[e:]
        return l, notes


# ---------------- 頁面 ----------------
def rulings_of(art, status_cfg):
    rs = list(status_cfg.get("rulings", []))
    for n in art["note"]:
        for r in re.findall(r"釋字第\s*(\d+)\s*號", n):
            if f"釋字第{r}號" not in rs: rs.append(f"釋字第{r}號")
        for y, r in re.findall(r"(\d+)\s*年\s*憲判字第\s*(\d+)\s*號", n):
            if f"{y}年憲判字第{r}號" not in rs: rs.append(f"{y}年憲判字第{r}號")
    return rs


def system_fields(ctx, no, idx):
    a, cfg = ctx.arts[no], ctx.cfg
    st = cfg.get("status", {}).get(no, {})
    if a["deleted"]: status, note = "已刪除", ""
    elif st: status, note = st["status"], st.get("note", "")
    elif a["note"]: status, note = "現行", "條文另附大法官解釋編註，見條文區"
    else: status, note = "現行", ""
    d = {"type": "法條", "law": ctx.full, "law_short": ctx.short, "article": no, "seq": idx + 1, "caption": a["caption"]}
    if not a["caption"] and not a["deleted"] and ctx.labels.get(no): d["caption_draft"] = ctx.labels[no]
    for unit in LEVEL:
        if unit in a["path"]: d[LEVEL_KEY[unit]] = a["path"][unit]
    d["status"] = status
    if note: d["status_note"] = note
    d["source"] = f"全國法規資料庫下載檔（下載日期{ctx.law['download']}；法規{ctx.law['revised_label']}{ctx.law['revised']}）"
    return d, status, note


def managed_block(ctx, no, idx, rulings, status, note):
    a, order = ctx.arts[no], ctx.order
    nav = []
    if idx > 0:
        p = order[idx - 1]; nav.append(ctx.link(p, f"← 第{p}條 {ctx.label(p, False)}".rstrip()))
    nav.append(f"[[{ctx.full}|法規總覽]]")
    if idx + 1 < len(order):
        n = order[idx + 1]; nav.append(ctx.link(n, (f"第{n}條 {ctx.label(n, False)}".rstrip()) + " →"))
    where = "／".join(a["path"][u] for u in LEVEL if u in a["path"])
    lines = [f"> [!quote] 條文　{where}".rstrip()]
    body = []
    for p in a["paras"]:
        l, notes = ctx.render(p, no)
        body += [l, ""]
        for kind, ttl, ls in notes:
            body.append(f"> [!note]{'-' if kind == 'note-' else ''} {ttl}")
            body += [x for t in ls for x in ("> " + t, ">")][:-1]
            body.append("")
    for ln in body[:-1]:
        lines.append(">" if not ln else "> " + ln)
    if a["note"]:
        lines += [">", "> > [!info]- 編註（全國法規資料庫附註）"]
        lines += [x for n in a["note"][1:] for x in ("> > " + n, "> >")][:-1]
    if rulings:
        lines += [">", "> 相關解釋／判決：" + "、".join(f"[[{r}]]" for r in rulings)]
    if note and status != "現行":
        lines += [">", f"> ⚠️ 效力：{status}。{note}"]
    return "\n".join([START, " ｜ ".join(nav), ""] + lines + ["", END])


TEMPLATE_TAIL = """## 法條解析


## 思考脈絡

%% 流程圖範本，要用時刪掉前後的百分號：
```mermaid
flowchart TD
    A[起點] --> B{判斷}
    B -->|是| C[結論一]
    B -->|否| D[結論二]
```
%%

## 實務見解


## 爭點


## 近年考試出現

| 年度 | 考試 | 科目 | 題號 | 考點 |
| --- | --- | --- | --- | --- |
|  |  |  |  |  |

## 相關考題


## 相關條文與解釋

"""


def build_page(ctx, no, idx, old_text, field_tag):
    a = ctx.arts[no]
    sysd, status, note = system_fields(ctx, no, idx)
    rulings = rulings_of(a, ctx.cfg.get("status", {}).get(no, {}))
    block = managed_block(ctx, no, idx, rulings, status, note)
    title = f"# {page_name(ctx.full, no)}" + (f" {a['caption']}" if a["caption"] else "")
    today = datetime.date.today().isoformat()
    if old_text is None:
        fm = dict(sysd)
        fm.update({"progress": "未整理", "exams": [], "related_rulings": [f"[[{r}]]" for r in rulings],
                   "tags": ["法條", field_tag, ctx.short],
                   "aliases": [f"{ctx.short}第{no}條", f"{ctx.short}{no}", f"{ctx.full}{no}"], "created": today})
        return dump_frontmatter(fm) + f"\n\n{title}\n\n{block}\n\n{TEMPLATE_TAIL}", "new"
    old_fm, rest = read_frontmatter(old_text)
    if START not in rest or END not in rest:
        return None, "no-markers"
    fm = {}
    for k in SYSTEM_KEYS:
        if k in sysd: fm[k] = sysd[k]
    for k, v in old_fm.items():
        if k not in SYSTEM_KEYS: fm[k] = v
    for k, v in USER_DEFAULTS.items(): fm.setdefault(k, v)
    rel = list(old_fm.get("related_rulings") or [])
    for r in rulings:
        if f"[[{r}]]" not in rel: rel.append(f"[[{r}]]")
    fm["related_rulings"] = rel
    pre, tail = rest.split(START, 1)
    tail = tail.split(END, 1)[1]
    pre = re.sub(r"^# .*$", title, pre, count=1, flags=re.M)
    new = dump_frontmatter(fm) + "\n" + pre + block + tail
    return new, ("same" if new == old_text else "updated")


def build_overview(ctx, field_tag):
    law = ctx.law
    fm = {"type": "法規總覽", "law": ctx.full, "law_short": ctx.short, "field": field_tag,
          "revised": law["revised"], "downloaded": law["download"], "articles": len(ctx.order),
          "tags": ["法規總覽", field_tag, ctx.short], "aliases": [ctx.short], "updated": datetime.date.today().isoformat()}
    out = [dump_frontmatter(fm), "", f"# {ctx.full}", "",
           "%% 本頁由腳本產生，每次執行會整頁覆寫，請勿在此寫筆記 %%", "",
           f"- 法規{law['revised_label']}：{law['revised']}", f"- 條文來源：全國法規資料庫（下載日期{law['download']}）",
           f"- 條數：{len(ctx.order)} 條（含已刪除條次）", ""]
    if law["history"]:
        out += ["> [!abstract]- 修正沿革"] + [x for h in law["history"] for x in ("> " + h, ">")][:-1] + [""]
    prev_a = False
    for item in law["toc"]:
        if item[0] == "h":
            if prev_a: out.append("")
            out += ["#" * item[1] + " " + item[2], ""]; prev_a = False
        else:
            no = item[1]
            out.append(f"- {ctx.link(no, '第' + no + '條')} {ctx.label(no, False)}".rstrip()); prev_a = True
    return "\n".join(out) + "\n"


def update_dashboard(full, category):
    if not os.path.exists(DASHBOARD): return "總覽頁不存在，略過"
    t = open(DASHBOARD, encoding="utf-8").read()
    cat = re.sub(r"^\d+-", "", category)
    # 使用者可能在 Obsidian 重排表格（欄位補空格對齊），比對時容許空白，寫回時保留第一欄原樣
    m = re.search(rf"^(\|\s*{re.escape(cat)}\s*\|)(.*)\|[ \t]*$", t, re.M)
    if not m: return f"總覽頁找不到「{cat}」列，請手動加入"
    if f"[[{full}]]" in m.group(2): return "總覽頁已有此法規"
    cell = m.group(2).strip()
    cell = f"[[{full}]]" if cell in ("（尚未建立）", "") else cell + f"、[[{full}]]"
    t = t[:m.start()] + f"{m.group(1)} {cell} |" + t[m.end():]
    open(DASHBOARD, "w", encoding="utf-8", newline="\n").write(t)
    return "總覽頁已加入此法規"


# ---------------- 核對 ----------------
def verify(ctx, outdir, report):
    import yaml
    names_all = set()
    for r, d, fs in os.walk(VAULT):
        names_all |= {os.path.splitext(f)[0] for f in fs} | set(fs)
    bad = []
    unresolved = {}
    for idx, no in enumerate(ctx.order):
        p = os.path.join(outdir, page_name(ctx.full, no) + ".md")
        if not os.path.exists(p): bad.append(f"缺頁面 第{no}條"); continue
        t = open(p, encoding="utf-8").read()
        try: read_frontmatter(t)
        except Exception as e: bad.append(f"YAML 錯誤 第{no}條 {e}")
        blk = t.split(START, 1)[1].split(END, 1)[0].split("> [!quote]", 1)[1]
        got = []
        for ln in blk.split("\n")[1:]:
            if not ln.startswith("> ") or ln.startswith("> >") or ln.startswith("> 相關解釋") or ln.startswith("> ⚠️"): continue
            ln = re.sub(r"\[\[[^\]|]+\|([^\]]+)\]\]", r"\1", ln[2:])
            got.append(re.sub(r" \((?:[^()]|\([^()]*\))*\)", "", ln))
        if got != ctx.arts[no]["paras"]: bad.append(f"條文還原不符 第{no}條")
        for tgt in re.findall(r"\[\[([^\]|#]+)", t):
            if tgt not in names_all: unresolved[tgt] = unresolved.get(tgt, 0) + 1
    ov = open(os.path.join(outdir, ctx.full + ".md"), encoding="utf-8").read()
    heads = set(re.findall(r"^#{2,6} (.+)$", ov, re.M))
    anchors = set()
    for f in os.listdir(outdir):
        anchors |= set(re.findall(rf"\[\[{re.escape(ctx.full)}#([^\]|]+)", open(os.path.join(outdir, f), encoding="utf-8").read()))
    for a in anchors - heads: bad.append(f"章節錨點找不到：{a}")
    report.append(f"核對：{len(ctx.order)} 條，問題 {len(bad)} 項")
    report += ["  " + b for b in bad[:30]]
    report.append("尚無頁面的連結目標（解釋／判決或尚未建立的法規）：" + (json.dumps(unresolved, ensure_ascii=False) if unresolved else "無"))
    return not bad


# ---------------- 子命令 ----------------
def cmd_scan(args):
    law = parse_docx(args.docx)
    full = args.law or law["title"]
    reg = registry()
    short = reg.get(full, {}).get("short", "")
    cfg_path = os.path.join(HERE, "設定", full + ".json")
    cfg = load_json(cfg_path, {"labels": {}, "external": {}, "status": {}, "law_names": {}})
    ctx = Ctx(law, full, short or full, cfg, reg)
    for idx, no in enumerate(ctx.order):
        for p in ctx.arts[no]["paras"]: ctx.render(p, no)
    pr = ctx.problems
    rep = [f"法規：{full}（檔內標題：{law['title']}）", f"{law['revised_label']}：{law['revised']}；下載日期：{law['download']}",
           f"條數：{len(ctx.order)}；層級：" + "、".join(sorted({u for a in ctx.arts.values() for u in a['path']}, key=LEVEL.get)),
           f"登錄狀態：{'已登錄' if full in reg else '尚未登錄（build 前要在 法規登錄.json 加一筆）'}", ""]
    rep.append(f"[1] 被引用但無官方條名、需自擬的條文（填 labels）：{sorted(pr['missing_label'], key=lambda x: [int(y) for y in x.split('-')])}")
    rep.append(f"[2] 無法判定法規名稱的引用（把正確法規名加進 law_names 或 external）：{len(pr['unknown_law'])} 處")
    rep += [f"    第{c}條：…{s}" for c, s in pr["unknown_law"]]
    rep.append(f"[3] 外部法規條文尚未提供、該法規也還沒建頁面（填 external，條文從全國法規資料庫取得）：")
    rep += [f"    第{c}條引用 {k}" for c, k in pr["missing_ext"]]
    rep.append("[4] 外部法規的條號範圍（在 ext_ranges 填「法規|起-迄」：範圍內實際存在的條號清單，依官網條文清單；"
               "引用舊版條號者填 {\"omit\": true, \"note\": …}）：")
    rep += [f"    第{c}條：{k}" for c, k in pr["ext_range"]]
    hist = [h for h in law["history"] if re.search(r"失其效力|違憲|憲判字|釋字", h)]
    rep.append("[5] 修正沿革中提到效力異動的紀錄（人工判讀後填 status）：")
    rep += ["    " + h for h in hist]
    notes = [no for no in ctx.order if ctx.arts[no]["note"]]
    rep.append(f"[6] 附編註的條文（腳本會自動抽出釋字／憲判字並建立連結）：{notes}")
    for k in pr["missing_label"]: cfg["labels"].setdefault(k, "")
    for c, k in pr["missing_ext"]:
        cfg["external"].setdefault(k, {"label": "", "pcode": "", "paras": []})
    save_json(cfg_path, cfg)
    rep.append(f"\n設定檔草稿已寫到：{cfg_path}（空白欄位要補）")
    print("\n".join(rep))


def load_ctx(args):
    reg = registry()
    if args.law not in reg: die(f"{args.law} 尚未在 法規登錄.json 登錄")
    law = parse_docx(args.docx)
    cfg = load_json(os.path.join(HERE, reg[args.law]["config"]), {})
    return Ctx(law, args.law, reg[args.law]["short"], cfg, reg), reg


def cmd_build(args):
    ctx, reg = load_ctx(args)
    cat = reg[args.law]["category"]
    if cat not in CATEGORIES: die(f"分類 {cat} 不在 {CATEGORIES}")
    field_tag = re.sub(r"^\d+-", "", cat)
    for k, v in ctx.labels.items():
        if not v: die(f"設定檔 labels 第{k}條還是空白")
    for k, v in ctx.ext.items():
        if v.get("omit"):
            if not v.get("note"): die(f"設定檔 external {k} 設了 omit 但沒有 note")
        elif not v.get("label") or not v.get("paras"): die(f"設定檔 external {k} 還沒填完")
    outdir = args.dry_run or law_dir(args.law, reg)
    pages, count = {}, {"new": 0, "updated": 0, "same": 0, "no-markers": 0}
    for idx, no in enumerate(ctx.order):
        p = os.path.join(outdir, page_name(args.law, no) + ".md")
        src = p if not args.dry_run else os.path.join(law_dir(args.law, reg), page_name(args.law, no) + ".md")
        old = open(src, encoding="utf-8").read() if os.path.exists(src) else None
        new, kind = build_page(ctx, no, idx, old, field_tag)
        count[kind] += 1
        if new is not None: pages[p] = new
    pr = ctx.problems
    if pr["unknown_law"] or pr["missing_ext"] or pr["missing_label"] or pr["ext_range"]:
        die("還有未處理的引用，先跑 scan 補設定檔：" + json.dumps(
            {k: (sorted(v) if isinstance(v, set) else v) for k, v in pr.items() if v}, ensure_ascii=False))
    os.makedirs(outdir, exist_ok=True)
    for p, t in pages.items():
        open(p, "w", encoding="utf-8", newline="\n").write(t)
    open(os.path.join(outdir, args.law + ".md"), "w", encoding="utf-8", newline="\n").write(build_overview(ctx, field_tag))
    gone = [f for f in os.listdir(outdir) if f.endswith(".md") and f != args.law + ".md"
            and f[:-3] not in {page_name(args.law, n) for n in ctx.order}]
    rep = [f"輸出位置：{os.path.abspath(outdir)}",
           f"新建 {count['new']}、更新 {count['updated']}、未變 {count['same']}、缺條文區標記而略過 {count['no-markers']}",
           "引用統計：" + json.dumps(ctx.stats, ensure_ascii=False),
           "新版法規已無、但資料夾裡還有的頁面（不會自動刪，請人工處理）：" + (", ".join(gone) or "無")]
    if not args.dry_run: rep.append(update_dashboard(args.law, cat))
    ok = verify(ctx, outdir, rep)
    print("\n".join(rep))
    if not ok: sys.exit(1)


def cmd_verify(args):
    ctx, reg = load_ctx(args)
    for no in ctx.order:
        for p in ctx.arts[no]["paras"]: ctx.render(p, no)
    rep = []
    ok = verify(ctx, law_dir(args.law, reg), rep)
    print("\n".join(rep)); sys.exit(0 if ok else 1)


def cmd_check_ext(args):
    reg = registry()
    cfg = load_json(os.path.join(HERE, reg[args.law]["config"]))
    allok = True
    for key, e in cfg.get("external", {}).items():
        law, no = key.split("|")
        if e.get("omit"): print(f"{key}：omit（不附條文，只附說明），略過比對"); continue
        if not e.get("pcode"): print(f"{key}：缺 pcode，無法比對"); allok = False; continue
        url = f"https://law.moj.gov.tw/LawClass/LawSingle.aspx?pcode={e['pcode']}&flno={no}"
        h = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=30).read().decode("utf-8")
        i0 = h.index('<div class="law-article">')
        blk = h[i0:h.index("clearfix", i0)]
        off = [html.unescape(re.sub(r"<[^>]+>", "", x)).strip() for x in re.findall(r'<div class="line-\d+[^"]*">(.*?)</div>', blk, re.S)]
        mine = [p.replace("　", "") for p in e["paras"]]
        same = off == mine
        allok &= same
        print(f"{key}：官方 {len(off)} 段、設定檔 {len(mine)} 段，{'逐段相同' if same else '不同'}")
        if not same:
            for a, b in zip(off, mine):
                if a != b: print(f"  官方：{a}\n  設定：{b}")
    sys.exit(0 if allok else 1)


def main():
    sys.stdout.reconfigure(encoding="utf-8")  # Windows 主控台預設 cp950，中文會亂碼
    sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("scan"); s.add_argument("--src", "--docx", dest="docx", required=True); s.add_argument("--law")
    b = sub.add_parser("build"); b.add_argument("--src", "--docx", dest="docx", required=True); b.add_argument("--law", required=True); b.add_argument("--dry-run")
    v = sub.add_parser("verify"); v.add_argument("--src", "--docx", dest="docx", required=True); v.add_argument("--law", required=True)
    c = sub.add_parser("check-ext"); c.add_argument("--law", required=True)
    a = ap.parse_args()
    {"scan": cmd_scan, "build": cmd_build, "verify": cmd_verify, "check-ext": cmd_check_ext}[a.cmd](a)


if __name__ == "__main__":
    main()
