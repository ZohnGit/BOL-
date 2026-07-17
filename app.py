import os
import queue
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import requests
import tkinter as tk
from requests.auth import HTTPBasicAuth
from tkinter import messagebox, ttk

from config import (
    BASE_API_URL,
    CLIENT_ID,
    CLIENT_SECRET,
    FULFILMENT_METHOD,
    LIST_ACCEPT_HEADER,
    LOGIN_URL,
    REQUEST_TIMEOUT,
    WAIT_SECONDS,
)


def find_first_datetime(value):
    if isinstance(value, str):
        if re.match(r"^\d{4}-\d{2}-\d{2}", value) or "T" in value:
            return value
        return None

    if isinstance(value, list):
        for item in value:
            found = find_first_datetime(item)
            if found:
                return found
        return None

    if isinstance(value, dict):
        for nested in value.values():
            found = find_first_datetime(nested)
            if found:
                return found
        return None

    return None


class BolApiClient:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({"Accept": LIST_ACCEPT_HEADER})
        self.access_token = None

    def login(self):
        response = self.session.post(
            LOGIN_URL,
            auth=HTTPBasicAuth(CLIENT_ID, CLIENT_SECRET),
            headers={
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data={"grant_type": "client_credentials"},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        payload = response.json()
        self.access_token = payload["access_token"]
        self.session.headers.update({"Authorization": f"Bearer {self.access_token}"})

    def fetch_shipments_page(self, page: int) -> List[Dict]:
        response = self.session.get(
            f"{BASE_API_URL}/shipments",
            params={"fulfilment-method": FULFILMENT_METHOD, "page": page},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        payload = response.json()
        return payload.get("shipments", [])

    def fetch_shipments_by_order(self, order_id: str) -> List[Dict]:
        response = self.session.get(
            f"{BASE_API_URL}/shipments",
            params={"order-id": order_id},
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        payload = response.json()
        return payload.get("shipments", [])

    def fetch_shipment_detail(self, shipment_id: str) -> Dict:
        response = self.session.get(
            f"{BASE_API_URL}/shipments/{shipment_id}",
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        return response.json()


class BolExporterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("BOL 已发货订单导出")
        self.root.geometry("760x560")
        self.root.minsize(700, 520)

        self.message_queue: "queue.Queue[tuple]" = queue.Queue()
        self.worker_running = False
        self.last_output_path: Optional[Path] = None

        self.page_var = tk.StringVar(value="1")
        self.status_var = tk.StringVar(value="等待开始")
        self.progress_text_var = tk.StringVar(value="0%")

        self._build_ui()
        self._poll_messages()

    def _build_ui(self):
        self.root.configure(bg="#f4efe5")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f4efe5")
        style.configure("TLabel", background="#f4efe5", foreground="#2f2419", font=("Microsoft YaHei UI", 11))
        style.configure("Header.TLabel", background="#f4efe5", foreground="#1f160d", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("TButton", font=("Microsoft YaHei UI", 10))
        style.configure("TEntry", font=("Consolas", 12))
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#eadfcd",
            background="#b56a2f",
            bordercolor="#eadfcd",
            lightcolor="#b56a2f",
            darkcolor="#b56a2f",
        )

        container = ttk.Frame(self.root, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="BOL 已发货订单导出工具", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="输入 page=N 时，程序会依次抓取第 1 页到第 N 页，并把结果合并后导出 Excel。",
            wraplength=680,
        ).pack(anchor="w", pady=(8, 18))

        form = ttk.Frame(container)
        form.pack(fill="x")

        ttk.Label(form, text="抓取页数 page").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(form, textvariable=self.page_var, width=12)
        entry.grid(row=0, column=1, padx=(12, 16), sticky="w")
        entry.focus()

        self.run_button = ttk.Button(form, text="开始抓取并导出", command=self.start_export)
        self.run_button.grid(row=0, column=2, sticky="w")

        self.open_button = ttk.Button(form, text="打开输出文件", command=self.open_output_file, state="disabled")
        self.open_button.grid(row=0, column=3, padx=(12, 0), sticky="w")

        status_frame = ttk.Frame(container)
        status_frame.pack(fill="x", pady=(24, 8))

        ttk.Label(status_frame, text="当前状态").pack(anchor="w")
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor="w", pady=(4, 0))

        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x", side="left", expand=True)
        ttk.Label(progress_frame, textvariable=self.progress_text_var, width=6).pack(side="left", padx=(10, 0))

        ttk.Label(container, text="运行日志").pack(anchor="w", pady=(24, 8))
        self.log_text = tk.Text(
            container,
            height=18,
            wrap="word",
            bg="#fffaf2",
            fg="#2f2419",
            relief="flat",
            font=("Consolas", 10),
            padx=12,
            pady=12,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

    def start_export(self):
        if self.worker_running:
            return

        raw_page = self.page_var.get().strip()
        if not raw_page.isdigit() or int(raw_page) <= 0:
            messagebox.showerror("参数错误", "请输入大于 0 的整数 page。")
            return

        if CLIENT_ID == "YOUR_BOL_CLIENT_ID" or CLIENT_SECRET == "YOUR_BOL_CLIENT_SECRET":
            messagebox.showerror("配置未完成", "请先在 config.py 中填写 CLIENT_ID 和 CLIENT_SECRET。")
            return

        page_limit = int(raw_page)
        self.worker_running = True
        self.last_output_path = None
        self.run_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self._set_progress(0, "准备开始")
        self._append_log(f"任务开始，目标抓取第 1 页到第 {page_limit} 页。")

        worker = threading.Thread(target=self._run_export, args=(page_limit,), daemon=True)
        worker.start()

    def _run_export(self, page_limit: int):
        try:
            client = BolApiClient()
            self.message_queue.put(("progress", 5, "正在登录 BOL"))
            self.message_queue.put(("log", "正在登录并获取 access token..."))
            client.login()

            all_shipments = []
            for page in range(1, page_limit + 1):
                progress = 5 + int((page / page_limit) * 35)
                self.message_queue.put(("progress", progress, f"正在抓取列表页 {page}/{page_limit}"))
                self.message_queue.put(("log", f"抓取列表页 {page}/{page_limit}"))
                shipments = client.fetch_shipments_page(page)
                self.message_queue.put(("log", f"第 {page} 页返回 {len(shipments)} 条 shipment。"))
                all_shipments.extend(shipments)

            parsed_orders = []
            seen_pairs = set()
            for shipment in all_shipments:
                shipment_id = shipment.get("shipmentId")
                order_id = shipment.get("order", {}).get("orderId")
                if not order_id:
                    continue
                pair = (shipment_id, order_id)
                if pair in seen_pairs:
                    continue
                seen_pairs.add(pair)
                parsed_orders.append({"shipmentId": shipment_id, "orderId": order_id})

            if not parsed_orders:
                raise RuntimeError("没有抓到任何订单数据，请检查 page 参数或接口返回。")

            self.message_queue.put(("log", f"累计得到 {len(parsed_orders)} 条待处理订单。"))

            result_rows = []
            total_orders = len(parsed_orders)
            for index, item in enumerate(parsed_orders, start=1):
                order_id = item["orderId"]
                shipment_list = client.fetch_shipments_by_order(order_id)
                shipment_id = item["shipmentId"]
                if shipment_list:
                    shipment_id = shipment_list[0].get("shipmentId") or shipment_id

                detail = client.fetch_shipment_detail(shipment_id)
                result_rows.append(
                    {
                        "orderId": detail.get("order", {}).get("orderId") or detail.get("orderId") or order_id,
                        "trackAndTrace": detail.get("transport", {}).get("trackAndTrace"),
                        "shipmentDateTime": find_first_datetime(detail),
                        "shipmentId": shipment_id,
                    }
                )

                progress = 40 + int((index / total_orders) * 55)
                self.message_queue.put(("progress", progress, f"正在处理订单 {index}/{total_orders}"))
                self.message_queue.put(("log", f"完成订单 {index}/{total_orders}: {order_id}"))
                time.sleep(WAIT_SECONDS)

            output_path = self._write_excel(result_rows)
            self.message_queue.put(("progress", 100, "导出完成"))
            self.message_queue.put(("done", str(output_path), f"已完成，共导出 {len(result_rows)} 条记录。"))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))
        finally:
            self.message_queue.put(("unlock",))

    def _write_excel(self, rows: List[Dict]) -> Path:
        output_dir = Path(__file__).resolve().parent / "exports"
        output_dir.mkdir(parents=True, exist_ok=True)
        filename = f"bol_shipments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = output_dir / filename

        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.drop_duplicates(subset=["orderId", "trackAndTrace", "shipmentId"], keep="first")
        df.to_excel(output_path, index=False)
        return output_path

    def _poll_messages(self):
        while True:
            try:
                message = self.message_queue.get_nowait()
            except queue.Empty:
                break

            kind = message[0]
            if kind == "progress":
                _, value, status = message
                self._set_progress(value, status)
            elif kind == "log":
                _, text = message
                self._append_log(text)
            elif kind == "done":
                _, path, status = message
                self.last_output_path = Path(path)
                self._append_log(f"导出文件已生成：{path}")
                self.status_var.set(status)
                self.open_button.configure(state="normal")
                messagebox.showinfo("完成", status)
            elif kind == "error":
                _, error_text = message
                self._append_log(f"发生错误：{error_text}")
                self.status_var.set("执行失败")
                messagebox.showerror("执行失败", error_text)
            elif kind == "unlock":
                self.worker_running = False
                self.run_button.configure(state="normal")

        self.root.after(150, self._poll_messages)

    def _set_progress(self, value: int, status: str):
        value = max(0, min(100, value))
        self.progress["value"] = value
        self.progress_text_var.set(f"{value}%")
        self.status_var.set(status)

    def _append_log(self, text: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {text}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def open_output_file(self):
        if not self.last_output_path or not self.last_output_path.exists():
            messagebox.showwarning("文件不存在", "当前没有可打开的输出文件。")
            return

        path = str(self.last_output_path)
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)


def main():
    root = tk.Tk()
    app = BolExporterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
