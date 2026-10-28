"""역할: TC 파일·대상 주소를 받아 실행·진행·결과·로그를 보여주는 프로그램 창입니다.

브라우저는 작업 스레드의 step_runner에서 조작하고, Tk 화면은 메인 스레드만
갱신합니다. 결과는 excel_io로 선택한 파일에 기록하며 pytest는 실행하지 않습니다.
"""

import json
import os
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urlsplit

from excel_io import read_cases, write_results
from step_runner import run_cases


PROJECT = Path(__file__).resolve().parent


class AutomationApp:
    """창의 상태와 작업 큐를 관리합니다. 브라우저 객체는 소유하지 않습니다."""

    def __init__(self, root):
        self.root = root
        self.running = False
        self.closing = False
        self.closed = False
        self.worker = None
        self.stop_event = Event()
        self.events = Queue()
        self.file_path = tk.StringVar(value=str(PROJECT / "SauceDemo_TC_실행용_v2.xlsx"))
        self.base_url = tk.StringVar(value="https://www.saucedemo.com/")
        self.operator = tk.StringVar(value="미기록")
        self.slow_view = tk.BooleanVar(value=True)
        self.current = tk.StringVar(value="대기 중")
        self.counts = tk.StringVar(value="Pass 0 / Fail 0 / N/A 0")
        self.save_state = tk.StringVar(value="결과는 선택한 TC 파일의 결과·비고에 저장됩니다.")
        root.title("SauceDemo 자동화 테스트")
        root.geometry("1100x700")
        root.minsize(850, 520)
        root.protocol("WM_DELETE_WINDOW", self.close)
        self._build()
        self.poll_id = root.after(80, self._poll)

    def _build(self):
        """합의한 상단 입력 영역과 세 개 탭을 만듭니다."""
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text="TC 파일 (.xlsx)").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.file_entry = ttk.Entry(top, textvariable=self.file_path)
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=4)
        self.browse_button = ttk.Button(top, text="파일 선택", command=self.browse)
        self.browse_button.grid(row=0, column=2, padx=(8, 0))
        ttk.Label(top, text="대상 주소").grid(row=1, column=0, sticky="w")
        self.url_entry = ttk.Entry(top, textvariable=self.base_url)
        self.url_entry.grid(row=1, column=1, columnspan=2, sticky="ew", pady=4)
        ttk.Label(top, text="실행자 / AI 지원").grid(row=2, column=0, sticky="w")
        self.operator_entry = ttk.Entry(top, textvariable=self.operator)
        self.operator_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=4)
        controls = ttk.Frame(top)
        controls.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.start_button = ttk.Button(controls, text="테스트 시작", command=self.start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="중지", command=self.stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        self.slow_check = ttk.Checkbutton(controls, text="천천히 보기 (5초 간격)", variable=self.slow_view)
        self.slow_check.pack(side="left", padx=8)
        ttk.Label(controls, textvariable=self.current).pack(side="left", padx=8)
        ttk.Label(controls, textvariable=self.counts).pack(side="right")
        ttk.Label(top, textvariable=self.save_state, wraplength=980).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.progress = self._table("TC 목록·진행", ["TC-ID", "시나리오", "진행 상태"], [90, 570, 220])
        self.results = self._table("실행 결과", ["TC-ID", "결과", "사유 / 증거"], [90, 90, 700])
        log_tab = ttk.Frame(self.tabs)
        self.tabs.add(log_tab, text="상세 로그")
        self.log = tk.Text(log_tab, wrap="word", state="disabled", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(log_tab, command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

    def _table(self, title, columns, widths):
        frame = ttk.Frame(self.tabs)
        self.tabs.add(frame, text=title)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for name, width in zip(columns, widths):
            tree.heading(name, text=name)
            tree.column(name, width=width, minwidth=70)
        scrollbar = ttk.Scrollbar(frame, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)
        for status, color in [("Pass", "#18733b"), ("Fail", "#bd2020"), ("N/A", "#806000")]:
            tree.tag_configure(status, foreground=color)
        return tree

    def browse(self):
        path = filedialog.askopenfilename(title="실행용 TC 선택", filetypes=[("Excel 통합 문서", "*.xlsx")])
        if path:
            self.file_path.set(path)

    def start(self):
        """입력 검사를 통과한 경우에만 작업 스레드를 하나 시작합니다."""
        if self.running or self.closing:
            return
        try:
            source = Path(self.file_path.get()).resolve()
            if not source.is_file() or source.suffix.lower() != ".xlsx":
                raise ValueError("존재하는 .xlsx TC 파일을 선택하세요.")
            url = self.base_url.get().strip()
            parsed = urlsplit(url)
            if parsed.scheme not in {"http", "https"} or not parsed.netloc:
                raise ValueError("대상 주소에 http:// 또는 https://와 호스트를 포함하세요.")
            cases = read_cases(source)
        except Exception as error:
            messagebox.showerror("실행 전 확인", str(error))
            return
        self.stop_event.clear()
        self.running = True
        self._controls(True)
        for tree in (self.progress, self.results):
            tree.delete(*tree.get_children())
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        for case in cases:
            self.progress.insert("", "end", iid=case["tc_id"], values=(case["tc_id"], case["scenario"], "대기"))
        self.counts.set("Pass 0 / Fail 0 / N/A 0")
        self.current.set(f"실행 준비: {len(cases)}개 TC")
        self.save_state.set("실행 중 · 엑셀 저장 전")
        self.worker = Thread(target=self._execute, args=(source, cases, url,
            self.operator.get().strip() or "미기록", 5000 if self.slow_view.get() else 0))
        self.worker.start()

    def _execute(self, source, cases, url, operator, slow_mo):
        """작업 스레드 전용: Tk 객체를 만지지 않고 이벤트만 큐에 넣습니다."""
        try:
            results = run_cases(cases, url, self.events.put, self.stop_event, slow_mo=slow_mo)
            try:
                write_results(source, results)
                saved, error = True, ""
                self.events.put({"type": "save_finished", "path": str(source)})
            except Exception as failure:
                saved, error = False, str(failure)
                self.events.put({"type": "save_failed", "message": error})
            # 실행자 정보는 엑셀 열을 추가하지 않고 해당 실행의 로그 폴더에 기록합니다.
            if results and results[0].get("evidence", {}).get("log"):
                summary = Path(results[0]["evidence"]["log"]).parent / "run-summary.json"
                summary.write_text(json.dumps({"operator": operator, "source": str(source),
                    "base_url": url, "slow_mo_ms": slow_mo, "excel_saved": saved, "save_error": error,
                    "results": results}, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as error:
            self.events.put({"type": "worker_error", "message": str(error)})
        finally:
            self.events.put({"type": "worker_finished"})

    def _controls(self, running):
        for widget in (self.start_button, self.browse_button, self.file_entry, self.url_entry, self.operator_entry, self.slow_check):
            widget.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")

    def stop(self):
        if self.running:
            self.stop_event.set()
            self.current.set("중지 요청됨 · 현재 동작 종료/제한시간 후 정리합니다.")
            self.stop_button.configure(state="disabled")

    def _poll(self):
        """메인 스레드에서 큐를 읽으므로 테스트 중에도 화면이 응답합니다."""
        try:
            while True:
                self._handle(self.events.get_nowait())
                if self.closed:
                    return
        except Empty:
            pass
        self.poll_id = self.root.after(80, self._poll)

    def _handle(self, event):
        kind = event["type"]
        self.log.configure(state="normal")
        self.log.insert("end", json.dumps(event, ensure_ascii=False) + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        tc_id = event.get("tc_id")
        if kind in {"tc_started", "step_started", "step_finished"}:
            state = "실행 중" if kind == "tc_started" else f"{event['step']}단계 {event['action']} · {'시작' if kind == 'step_started' else '완료'}"
            self.current.set(f"{tc_id}: {state}")
            self.progress.set(tc_id, "진행 상태", state)
            self.progress.selection_set(tc_id)
            self.progress.see(tc_id)
        elif kind == "tc_finished":
            status = event["status"]
            self.progress.set(tc_id, "진행 상태", status)
            self.progress.item(tc_id, tags=(status,))
            details = event.get("reason", "")
            details += " " + " / ".join(event.get("evidence", {}).values())
            if self.results.exists(tc_id):
                self.results.delete(tc_id)
            self.results.insert("", "end", iid=tc_id, values=(tc_id, status, details.strip()), tags=(status,))
            statuses = [self.results.item(item, "values")[1] for item in self.results.get_children()]
            self.counts.set(" / ".join(f"{value} {statuses.count(value)}" for value in ["Pass", "Fail", "N/A"]))
        elif kind == "save_finished":
            self.save_state.set(f"엑셀 저장 완료: {event['path']}")
        elif kind == "save_failed":
            self.save_state.set("엑셀 저장 실패 · 실행 결과는 유지됩니다. 상세 로그와 실행 폴더를 확인하세요.")
            messagebox.showerror("엑셀 저장 실패", event["message"] + "\n파일을 닫고 다시 실행하세요. 이번 실행 결과는 로그에 남습니다.")
        elif kind == "worker_error":
            self.save_state.set("작업 오류 · 저장 완료 여부와 상세 로그를 확인하세요.")
            messagebox.showerror("작업 오류", event["message"])
        elif kind == "worker_finished":
            self.running = False
            self._controls(False)
            self.current.set("종료 · 실행 결과와 저장 상태를 확인하세요.")
            if self.closing:
                self.closed = True
                self.root.destroy()

    def close(self):
        """실행 중에는 중지·저장·브라우저 정리를 마친 뒤 닫습니다."""
        if self.running:
            self.closing = True
            self.stop()
        else:
            self.closed = True
            self.root.after_cancel(self.poll_id)
            self.root.destroy()


def main():
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(PROJECT / ".browsers"))
    root = tk.Tk()
    AutomationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
