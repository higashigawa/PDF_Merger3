#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF結合ツール (pdf_merger.py)

- ファイル選択ダイアログ / ドラッグ＆ドロップ の両方でPDFを追加
- 一覧上でドラッグまたは ↑↓ ボタンで結合順を並べ替え
- ファイルごとにページ範囲を指定 (1-3,5 / 5- / -3 / 3,1 / 5-3 など)
- 同じPDFを別のページ範囲で複数回使える「複製」
- 各ファイル名をしおり(アウトライン)として結合先に付与
- パスワード付きPDFは追加時にパスワードを入力

依存:
    pip install pypdf
    pip install tkinterdnd2   # ドラッグ＆ドロップ用（無い場合はD&Dのみ無効化して動作）

Windows / Ubuntu 両対応。
"""

import os
import re
import sys
import datetime
import threading
import traceback
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# ---------------------------------------------------------------- 依存モジュール

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:  # pragma: no cover
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "pypdf が見つかりません",
        "PDFの読み書きに pypdf が必要です。\n\n"
        "  pip install pypdf\n\n"
        "を実行してから起動してください。",
    )
    sys.exit(1)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES

    DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None
    DND_FILES = None
    DND_AVAILABLE = False


def app_dir() -> str:
    """実行ファイル(PyInstaller)またはスクリプトのあるディレクトリ。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------- データモデル


_TOKEN_RE = re.compile(r"^(\d*)(-?)(\d*)$")


def normalize_spec(spec: str) -> str:
    """全角数字や空白、読点を吸収して比較・保存しやすい形にそろえる。"""
    text = unicodedata.normalize("NFKC", spec or "")
    text = text.replace("、", ",").replace("，", ",").replace("〜", "-").replace("–", "-")
    return "".join(text.split())


def parse_page_range(spec: str, total: int) -> list[int]:
    """
    "1-3,5" 形式の指定を 0 始まりのページ番号リストに変換する。

    空文字      -> 全ページ
    "1-3,5"     -> 1,2,3,5 ページ目
    "5-"        -> 5ページ目から最後まで
    "-3"        -> 先頭から3ページ目まで
    "3,1"       -> 指定した順に並べ替え（重複指定も可）
    "5-3"       -> 逆順に取り込み

    解釈できない場合は ValueError（日本語メッセージ）を投げる。
    """
    text = normalize_spec(spec)
    if not text:
        return list(range(total))

    pages: list[int] = []
    for token in text.split(","):
        if not token:
            continue
        mo = _TOKEN_RE.match(token)
        if not mo:
            raise ValueError(f"「{token}」は解釈できません")
        start_s, dash, end_s = mo.groups()
        if dash:
            start = int(start_s) if start_s else 1
            end = int(end_s) if end_s else total
        else:
            if not start_s:
                raise ValueError(f"「{token}」は解釈できません")
            start = end = int(start_s)
        for n in (start, end):
            if not 1 <= n <= total:
                raise ValueError(f"ページ {n} は範囲外です（1〜{total}）")
        step = 1 if end >= start else -1
        pages.extend(range(start, end + step, step))

    if not pages:
        raise ValueError("ページが指定されていません")
    return [n - 1 for n in pages]


class PdfEntry:
    """結合対象の1ファイル分の情報。"""

    def __init__(self, path: str, pages: int, password: str = "", page_spec: str = ""):
        self.path = path
        self.pages = pages
        self.password = password
        self.page_spec = page_spec

    @property
    def name(self) -> str:
        return os.path.basename(self.path)

    @property
    def key(self) -> str:
        return os.path.normcase(os.path.abspath(self.path))

    @property
    def spec_label(self) -> str:
        return self.page_spec if self.page_spec else "全ページ"

    def resolved(self, total: int | None = None) -> list[int]:
        return parse_page_range(self.page_spec, self.pages if total is None else total)

    @property
    def out_pages(self) -> int:
        """出力されるページ数。指定が不正なら -1。"""
        try:
            return len(self.resolved())
        except ValueError:
            return -1


# ---------------------------------------------------------------- ダイアログ


class PageRangeDialog(simpledialog.Dialog):
    """ページ範囲を入力させ、OK時に検証する小さなモーダルダイアログ。"""

    HELP = (
        "例:\n"
        "  1-3,5     1〜3ページ目と5ページ目\n"
        "  5-        5ページ目から最後まで\n"
        "  -3        先頭から3ページ目まで\n"
        "  3,1       この順に並べ替える（重複指定も可）\n"
        "  5-3       逆順に取り込む\n"
        "空欄にすると全ページを対象にします。"
    )

    def __init__(self, parent, initial: str, max_pages: int, count: int = 1):
        self.initial = initial
        self.max_pages = max_pages
        self.count = count
        self.result_spec: str | None = None
        super().__init__(parent, "ページ範囲の指定")

    def body(self, master):
        master.columnconfigure(0, weight=1)
        head = (
            f"結合するページを指定してください（全{self.max_pages}ページ）"
            if self.count == 1
            else f"選択中の{self.count}ファイルに同じ指定を適用します"
            f"（最少ファイルは{self.max_pages}ページ）"
        )
        ttk.Label(master, text=head).grid(row=0, column=0, sticky="w")

        self.var = tk.StringVar(value=self.initial)
        self.entry = ttk.Entry(master, textvariable=self.var, width=40)
        self.entry.grid(row=1, column=0, sticky="ew", pady=(6, 8))
        self.entry.select_range(0, "end")

        ttk.Label(master, text=self.HELP, foreground="#555", justify="left").grid(
            row=2, column=0, sticky="w"
        )
        return self.entry

    def validate(self) -> bool:
        try:
            parse_page_range(self.var.get(), self.max_pages)
        except ValueError as exc:
            messagebox.showwarning("指定を確認してください", str(exc), parent=self)
            return False
        return True

    def apply(self) -> None:
        self.result_spec = normalize_spec(self.var.get())


# ---------------------------------------------------------------- アプリ本体


class PdfMergerApp:
    PAD = 8

    def __init__(self, root: tk.Tk):
        self.root = root
        self.entries: list[PdfEntry] = []
        self._drag_index: int | None = None
        self._merging = False

        root.title("PDF結合ツール")
        root.geometry("760x520")
        root.minsize(620, 420)

        self._build_ui()
        self._setup_dnd()
        self._refresh()

    # ------------------------------------------------------------ UI構築

    def _build_ui(self) -> None:
        root = self.root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        # 上部: 追加系ボタン
        top = ttk.Frame(root, padding=(self.PAD, self.PAD, self.PAD, 0))
        top.grid(row=0, column=0, sticky="ew")
        ttk.Button(top, text="ファイルを追加...", command=self.add_files).pack(
            side="left"
        )
        ttk.Button(top, text="フォルダを追加...", command=self.add_folder).pack(
            side="left", padx=(6, 0)
        )
        ttk.Button(top, text="すべてクリア", command=self.clear_all).pack(
            side="left", padx=(6, 0)
        )

        hint = (
            "ここにPDFをドラッグ＆ドロップできます"
            if DND_AVAILABLE
            else "D&Dを使うには tkinterdnd2 を入れてください"
        )
        ttk.Label(top, text=hint, foreground="#666").pack(side="right")

        # 中央: 一覧 + 並べ替えボタン
        center = ttk.Frame(root, padding=self.PAD)
        center.grid(row=1, column=0, sticky="nsew")
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=1)

        columns = ("no", "name", "pages", "spec", "out", "path")
        self.tree = ttk.Treeview(
            center, columns=columns, show="headings", selectmode="extended"
        )
        for col, text, width, anchor in (
            ("no", "#", 36, "center"),
            ("name", "ファイル名", 200, "w"),
            ("pages", "総ページ", 64, "center"),
            ("spec", "ページ範囲", 110, "w"),
            ("out", "出力", 48, "center"),
            ("path", "場所", 220, "w"),
        ):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=anchor)
        self.tree.tag_configure("invalid", foreground="#c00000")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(center, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)

        side = ttk.Frame(center)
        side.grid(row=0, column=2, sticky="ns", padx=(self.PAD, 0))
        ttk.Button(side, text="↑", width=4, command=lambda: self.move(-1)).pack()
        ttk.Button(side, text="↓", width=4, command=lambda: self.move(1)).pack(
            pady=(4, 0)
        )
        ttk.Button(side, text="範囲...", width=6, command=self.edit_range).pack(
            pady=(12, 0)
        )
        ttk.Button(side, text="複製", width=6, command=self.duplicate_selected).pack(
            pady=(4, 0)
        )
        ttk.Button(side, text="削除", width=6, command=self.remove_selected).pack(
            pady=(12, 0)
        )
        ttk.Button(side, text="名前順", width=6, command=self.sort_by_name).pack(
            pady=(4, 0)
        )

        # 一覧内ドラッグで並べ替え
        self.tree.bind("<ButtonPress-1>", self._on_row_press)
        self.tree.bind("<B1-Motion>", self._on_row_drag)
        self.tree.bind("<ButtonRelease-1>", self._on_row_release)
        self.tree.bind("<Double-1>", self.edit_range)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # 下部: しおり設定と実行
        bottom = ttk.Frame(root, padding=(self.PAD, 0, self.PAD, self.PAD))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        self.bookmark_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            bottom,
            text="ファイル名をしおりとして追加する",
            variable=self.bookmark_var,
        ).grid(row=0, column=0, sticky="w")

        self.merge_btn = ttk.Button(
            bottom, text="結合して保存...", command=self.merge
        )
        self.merge_btn.grid(row=0, column=1, sticky="e")

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        self.status = tk.StringVar(value="PDFを追加してください。")
        ttk.Label(root, textvariable=self.status, anchor="w", relief="sunken").grid(
            row=3, column=0, sticky="ew"
        )

    def _setup_dnd(self) -> None:
        if not DND_AVAILABLE:
            return
        for widget in (self.tree, self.root):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)

    # ------------------------------------------------------------ 追加処理

    def _on_drop(self, event) -> None:
        # "{C:/path with space/a.pdf} C:/b.pdf" 形式を安全に分解
        paths = self.root.tk.splitlist(event.data)
        self._add_paths(paths)

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="結合するPDFを選択",
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")],
        )
        if paths:
            self._add_paths(paths)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="PDFの入ったフォルダを選択")
        if folder:
            self._add_paths([folder])

    def _collect_pdfs(self, paths) -> list[str]:
        found: list[str] = []
        for p in paths:
            if os.path.isdir(p):
                for dirpath, _dirnames, filenames in os.walk(p):
                    for fn in sorted(filenames):
                        if fn.lower().endswith(".pdf"):
                            found.append(os.path.join(dirpath, fn))
            elif p.lower().endswith(".pdf"):
                found.append(p)
        return found

    def _add_paths(self, paths) -> None:
        if self._merging:
            return
        candidates = self._collect_pdfs(paths)
        if not candidates:
            self.status.set("PDFファイルが見つかりませんでした。")
            return

        known = {e.key for e in self.entries}
        added = skipped = 0
        errors: list[str] = []

        for path in candidates:
            key = os.path.normcase(os.path.abspath(path))
            if key in known:
                skipped += 1
                continue
            entry = self._make_entry(path, errors)
            if entry is None:
                continue
            self.entries.append(entry)
            known.add(key)
            added += 1

        self._refresh()
        msg = f"{added}件を追加しました。"
        if skipped:
            msg += f" 重複{skipped}件をスキップ（同じPDFを再利用するなら「複製」）。"
        if errors:
            msg += f" 読み込み失敗{len(errors)}件。"
            messagebox.showwarning("読み込めないファイル", "\n".join(errors[:10]))
        self.status.set(msg)

    def _make_entry(self, path: str, errors: list[str]) -> PdfEntry | None:
        """ページ数を読み、必要ならパスワードを尋ねてエントリを作る。"""
        try:
            reader = PdfReader(path)
            password = ""
            if reader.is_encrypted:
                if not reader.decrypt(""):
                    password = simpledialog.askstring(
                        "パスワード",
                        f"{os.path.basename(path)} のパスワードを入力してください:",
                        show="*",
                        parent=self.root,
                    )
                    if password is None:
                        return None
                    if not reader.decrypt(password):
                        errors.append(f"{os.path.basename(path)}: パスワードが違います")
                        return None
            return PdfEntry(path, len(reader.pages), password)
        except Exception as exc:  # 壊れたPDFなど
            errors.append(f"{os.path.basename(path)}: {exc}")
            return None

    # ------------------------------------------------------------ 一覧操作

    def _refresh(self) -> None:
        selected_keys = {
            self.entries[i].key for i in self._selected_indices() if i < len(self.entries)
        }
        self.tree.delete(*self.tree.get_children())
        invalid = 0
        total_pages = 0
        for i, entry in enumerate(self.entries):
            out = entry.out_pages
            if out < 0:
                invalid += 1
            else:
                total_pages += out
            iid = self.tree.insert(
                "",
                "end",
                values=(
                    i + 1,
                    entry.name,
                    entry.pages,
                    entry.spec_label,
                    out if out >= 0 else "エラー",
                    os.path.dirname(entry.path),
                ),
                tags=("invalid",) if out < 0 else (),
            )
            if entry.key in selected_keys:
                self.tree.selection_add(iid)

        ready = bool(self.entries) and not invalid and not self._merging
        self.merge_btn.state(["!disabled"] if ready else ["disabled"])
        if invalid:
            self.status.set(f"{invalid}件のページ範囲が不正です。修正してください。")
        elif self.entries:
            self.status.set(
                f"{len(self.entries)}ファイル / 出力{total_pages}ページ"
            )

    def _selected_indices(self) -> list[int]:
        children = self.tree.get_children()
        return sorted(children.index(iid) for iid in self.tree.selection())

    def _select_indices(self, indices) -> None:
        children = self.tree.get_children()
        self.tree.selection_set([children[i] for i in indices if 0 <= i < len(children)])

    def move(self, delta: int) -> None:
        indices = self._selected_indices()
        if not indices:
            return
        order = indices if delta < 0 else reversed(indices)
        moved: list[int] = []
        for i in order:
            j = i + delta
            if j < 0 or j >= len(self.entries) or j in moved:
                moved.append(i)
                continue
            self.entries[i], self.entries[j] = self.entries[j], self.entries[i]
            moved.append(j)
        self._refresh()
        self._select_indices(moved)

    def edit_range(self, event=None) -> None:
        """選択行のページ範囲を編集する。複数選択時は同じ指定をまとめて適用。"""
        if self._merging:
            return
        if event is not None and not self.tree.identify_row(event.y):
            return
        indices = self._selected_indices()
        if not indices:
            return
        targets = [self.entries[i] for i in indices]
        limit = min(e.pages for e in targets)
        dialog = PageRangeDialog(
            self.root, targets[0].page_spec, limit, count=len(targets)
        )
        if dialog.result_spec is None:
            return
        for entry in targets:
            entry.page_spec = dialog.result_spec
        self._refresh()
        self._select_indices(indices)
        label = dialog.result_spec or "全ページ"
        self.status.set(f"{len(targets)}件のページ範囲を「{label}」にしました。")

    def duplicate_selected(self) -> None:
        """同じPDFを別のページ範囲でも使えるように選択行を複製する。"""
        indices = set(self._selected_indices())
        if not indices:
            return
        new_entries: list[PdfEntry] = []
        new_sel: list[int] = []
        for i, entry in enumerate(self.entries):
            new_entries.append(entry)
            if i in indices:
                new_entries.append(
                    PdfEntry(entry.path, entry.pages, entry.password, entry.page_spec)
                )
                new_sel.append(len(new_entries) - 1)
        self.entries = new_entries
        self._refresh()
        self._select_indices(new_sel)
        self.status.set(f"{len(indices)}件を複製しました。ページ範囲を指定できます。")

    def remove_selected(self) -> None:
        indices = set(self._selected_indices())
        if not indices:
            return
        self.entries = [e for i, e in enumerate(self.entries) if i not in indices]
        self._refresh()
        self.status.set(f"{len(indices)}件を削除しました。")

    def sort_by_name(self) -> None:
        self.entries.sort(key=lambda e: e.name.lower())
        self._refresh()

    def clear_all(self) -> None:
        self.entries.clear()
        self._refresh()
        self.status.set("一覧を空にしました。")

    # --- 一覧内のドラッグ並べ替え

    def _on_row_press(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        self._drag_index = (
            self.tree.get_children().index(iid) if iid else None
        )

    def _on_row_drag(self, event) -> None:
        if self._drag_index is None:
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        target = self.tree.get_children().index(iid)
        if target == self._drag_index:
            return
        entry = self.entries.pop(self._drag_index)
        self.entries.insert(target, entry)
        self._drag_index = target
        self._refresh()
        self._select_indices([target])

    def _on_row_release(self, _event) -> None:
        self._drag_index = None

    # ------------------------------------------------------------ 結合

    def merge(self) -> None:
        if self._merging or not self.entries:
            return

        default_dir = os.path.dirname(self.entries[0].path) or app_dir()
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = filedialog.asksaveasfilename(
            title="結合後のPDFの保存先",
            initialdir=default_dir,
            initialfile=f"merged_{stamp}.pdf",
            defaultextension=".pdf",
            filetypes=[("PDFファイル", "*.pdf")],
        )
        if not out_path:
            return

        sources = {e.key for e in self.entries}
        if os.path.normcase(os.path.abspath(out_path)) in sources:
            messagebox.showerror(
                "保存先が不正",
                "結合元のファイルを保存先に指定できません。別の名前にしてください。",
            )
            return

        self._merging = True
        self.merge_btn.state(["disabled"])
        self.progress.configure(maximum=len(self.entries), value=0)
        self.status.set("結合中...")

        entries = list(self.entries)
        add_bookmarks = self.bookmark_var.get()
        thread = threading.Thread(
            target=self._merge_worker,
            args=(entries, out_path, add_bookmarks),
            daemon=True,
        )
        thread.start()

    def _merge_worker(self, entries, out_path, add_bookmarks) -> None:
        try:
            writer = PdfWriter()
            for done, entry in enumerate(entries, start=1):
                reader = PdfReader(entry.path)
                if reader.is_encrypted:
                    reader.decrypt(entry.password)
                # 追加時からファイルが差し替わっている場合に備え、実ページ数で再解釈
                try:
                    indices = entry.resolved(len(reader.pages))
                except ValueError as exc:
                    raise ValueError(f"{entry.name}: {exc}") from None
                start = len(writer.pages)
                for idx in indices:
                    writer.add_page(reader.pages[idx])
                if add_bookmarks:
                    title = os.path.splitext(entry.name)[0]
                    if entry.page_spec:
                        title += f" ({entry.page_spec})"
                    writer.add_outline_item(title, start)
                self.root.after(0, self.progress.configure, {"value": done})

            total = len(writer.pages)
            with open(out_path, "wb") as fp:
                writer.write(fp)
            writer.close()
            self.root.after(0, self._merge_done, out_path, total, None)
        except Exception:
            self.root.after(0, self._merge_done, out_path, 0, traceback.format_exc())

    def _merge_done(self, out_path, total_pages, error) -> None:
        self._merging = False
        self.progress.configure(value=0)
        self._refresh()
        if error:
            self.status.set("結合に失敗しました。")
            messagebox.showerror("結合エラー", error)
            return
        self.status.set(f"保存しました: {out_path} ({total_pages}ページ)")
        if messagebox.askyesno(
            "完了",
            f"{total_pages}ページのPDFを保存しました。\n\n{out_path}\n\n"
            "保存先のフォルダを開きますか？",
        ):
            self._open_folder(os.path.dirname(out_path))

    @staticmethod
    def _open_folder(folder: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}" >/dev/null 2>&1 &')
        except Exception:
            pass


def main() -> None:
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform.startswith("win") else "clam")
    except tk.TclError:
        pass
    PdfMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
