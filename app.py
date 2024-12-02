import re
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Optional

import requests


def setup_database(db_name: str = "contacts.db") -> None:
    """
    初始化 SQLite 資料庫，創建 contacts 表格。

    :param db_name: 資料庫名稱，預設為 contacts.db
    """
    try:
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS contacts (
                iid INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                title TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE
            )
        """)
        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror("資料庫錯誤", f"初始化資料庫時發生錯誤: {e}")
    finally:
        conn.close()


def save_to_database(contacts: List[Dict[str, str]], db_name: str = "contacts.db") -> None:
    """
    將聯絡資訊存入 SQLite 資料庫，避免重複記錄。

    :param contacts: 聯絡資訊列表，每個聯絡人為一個字典
    :param db_name: 資料庫名稱，預設為 contacts.db
    """
    try:
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()
        for contact in contacts:
            try:
                cursor.execute("""
                    INSERT INTO contacts (name, title, email)
                    VALUES (?, ?, ?)
                """, (contact['name'], contact['title'], contact['email']))
            except sqlite3.IntegrityError:
                # 忽略重複的 email
                continue
        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror("資料庫錯誤", f"存入資料庫時發生錯誤: {e}")
    finally:
        conn.close()


def parse_contacts(html: str) -> List[Dict[str, str]]:
    """
    使用正規表達式解析 HTML 內容，提取聯絡資訊。

    :param html: 網頁的 HTML 內容
    :return: 聯絡資訊列表，每個聯絡人為一個字典
    """
    contacts = []
    # 假設每位教師的資訊在特定的 HTML 結構中，例如每個教師在 <div class="teacher"> 中
    # 需要根據實際 HTML 結構調整正則表達式
    teacher_pattern = re.compile(
        r'<div\s+class=["\']teacher["\'].*?>.*?'
        r'<p\s+class=["\']name["\'].*?>(?P<name>.*?)</p>.*?'
        r'<p\s+class=["\']title["\'].*?>(?P<title>.*?)</p>.*?'
        r'<a\s+href=["\']mailto:(?P<email>[^"\']+)["\'].*?>(?P=email)</a>.*?</div>',
        re.DOTALL
    )
    matches = teacher_pattern.finditer(html)
    for match in matches:
        name = match.group('name').strip()
        title = match.group('title').strip()
        email = match.group('email').strip()
        contacts.append({
            'name': name if name else 'N/A',
            'title': title if title else 'N/A',
            'email': email if email else 'N/A'
        })

    # 如果上述模式未匹配到任何資料，嘗試另一種模式
    if not contacts:
        # 例如，教師資訊可能在表格中
        table_pattern = re.compile(
            r'<tr.*?>\s*<td.*?>(?P<name>.*?)</td>\s*<td.*?>(?P<title>.*?)</td>\s*<td.*?>.*?mailto:(?P<email>[^"\']+)["\'].*?</td>',
            re.DOTALL
        )
        matches = table_pattern.finditer(html)
        for match in matches:
            name = match.group('name').strip()
            title = match.group('title').strip()
            email = match.group('email').strip()
            contacts.append({
                'name': name if name else 'N/A',
                'title': title if title else 'N/A',
                'email': email if email else 'N/A'
            })

    return contacts


def scrape_contacts(url: str) -> Optional[List[Dict[str, str]]]:
    """
    從指定的 URL 抓取聯絡資訊。

    :param url: 目標網頁的 URL
    :return: 聯絡資訊列表，若失敗則回傳 None
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        contacts = parse_contacts(response.text)
        return contacts
    except requests.exceptions.HTTPError as http_err:
        messagebox.showerror("HTTP錯誤", f"HTTP 錯誤發生: {http_err}")
    except requests.exceptions.ConnectionError:
        messagebox.showerror("連線錯誤", "無法連接到網路。")
    except requests.exceptions.Timeout:
        messagebox.showerror("逾時錯誤", "請求超時。")
    except requests.exceptions.RequestException as e:
        messagebox.showerror("請求錯誤", f"請求時發生錯誤: {e}")
    return None


def display_contacts(contacts: List[Dict[str, str]], tree: ttk.Treeview) -> None:
    """
    在 Tkinter 的 Treeview 中顯示聯絡資訊。

    :param contacts: 聯絡資訊列表
    :param tree: Tkinter 的 Treeview 元件
    """
    for contact in contacts:
        tree.insert('', tk.END, values=(contact['name'], contact['title'], contact['email']))


class ContactApp:
    """
    聯絡資訊抓取與顯示應用程式。
    """

    def __init__(self, root: tk.Tk) -> None:
        """
        初始化應用程式。

        :param root: Tkinter 的根視窗
        """
        self.root = root
        self.root.title("聯絡資訊抓取器")
        self.root.geometry("640x480")
        self.url = tk.StringVar(value="https://ai.ncut.edu.tw/app/index.php?Action=mobileloadmod&Type=mobile_rcg_mstr&Nbr=730")
        setup_database()

        self.create_widgets()

    def create_widgets(self) -> None:
        """
        創建並佈局所有的元件。
        """
        # URL 輸入框
        url_label = ttk.Label(self.root, text="URL:")
        url_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')

        url_entry = ttk.Entry(self.root, textvariable=self.url)
        url_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky='ew')

        # 抓取按鈕
        fetch_button = ttk.Button(self.root, text="抓取", command=self.fetch_contacts)
        fetch_button.grid(row=0, column=3, padx=5, pady=5, sticky='e')

        # Treeview 顯示聯絡資訊
        columns = ("name", "title", "email")
        self.tree = ttk.Treeview(self.root, columns=columns, show='headings')
        self.tree.heading("name", text="姓名")
        self.tree.heading("title", text="分機")
        self.tree.heading("email", text="Email")

        self.tree.column("name", anchor='center')
        self.tree.column("title", anchor='center')
        self.tree.column("email", anchor='center')

        self.tree.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='nsew')

        # 設定列與欄的權重，以便調整視窗大小時自動調整
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

    def fetch_contacts(self) -> None:
        """
        按下抓取按鈕後的處理流程：抓取、顯示並存入資料庫。
        """
        url = self.url.get().strip()
        if not url:
            messagebox.showwarning("輸入錯誤", "請輸入有效的 URL。")
            return

        contacts = scrape_contacts(url)
        if contacts is not None:
            if contacts:
                # 清除現有的 Treeview 資料
                for item in self.tree.get_children():
                    self.tree.delete(item)
                # 顯示新的聯絡資訊
                display_contacts(contacts, self.tree)
                # 儲存至資料庫
                save_to_database(contacts)
                messagebox.showinfo("成功", "聯絡資訊已成功抓取並存入資料庫。")
            else:
                messagebox.showinfo("結果", "未抓取到任何聯絡資訊。")


def main() -> None:
    """
    主函式，啟動應用程式。
    """
    root = tk.Tk()
    app = ContactApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
