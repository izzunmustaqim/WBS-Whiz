import PyPDF2
import requests
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext

class App:
    def __init__(self, master):
        self.master = master
        master.title("AI API クライアント")

        self.pdf_label = tk.Label(master, text="PDFファイルを選択:")
        self.pdf_label.pack()

        self.pdf_button = tk.Button(master, text="参照", command=self.browse_pdf)
        self.pdf_button.pack()

        self.prompt_label = tk.Label(master, text="プロンプト:")
        self.prompt_label.pack()

        self.prompt_entry = scrolledtext.ScrolledText(master, wrap=tk.WORD, height=5)
        self.prompt_entry.pack()

        self.api_key_label = tk.Label(master, text="APIキー:")
        self.api_key_label.pack()

        self.api_key_entry = tk.Entry(master, show="*")  # APIキーを隠す
        self.api_key_entry.pack()

        self.send_button = tk.Button(master, text="送信", command=self.send_request)
        self.send_button.pack()

        self.response_label = tk.Label(master, text="応答:")
        self.response_label.pack()

        self.response_text = scrolledtext.ScrolledText(master, wrap=tk.WORD, height=10)
        self.response_text.pack()

        self.pdf_path = None


    def browse_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            self.pdf_label.config(text=f"選択されたPDF: {self.pdf_path}")


    def send_request(self):
        if not self.pdf_path:
            self.response_text.insert(tk.END, "PDFファイルを選択してください\n")
            return

        try:
            with open(self.pdf_path, "rb") as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                text = ""
                for page in range(len(pdf_reader.pages)):
                    text += pdf_reader.pages[page].extract_text()

            prompt = self.prompt_entry.get("1.0", tk.END).strip()
            api_key = self.api_key_entry.get()
            combined_text = text + "\n" + prompt

            headers = {
                "Content-Type": "application/json",
                "api-key": api_key
            }
            data = {"messages": [{"role": "user", "content": combined_text}]}

            response = requests.post("https://ai-foundation-api.app/ai-foundation/chat-ai/gpt4", headers=headers, json=data)
            response.raise_for_status() # HTTPエラーを例外として送出

            self.response_text.insert(tk.END, response.json())

        except FileNotFoundError:
            self.response_text.insert(tk.END, "PDFファイルが見つかりません\n")
        except requests.exceptions.RequestException as e:
            self.response_text.insert(tk.END, f"APIリクエストエラー: {e}\n")
        except PyPDF2.errors.PdfReadError:
            self.response_text.insert(tk.END, "PDFファイルの読み込みに失敗しました\n")
        except Exception as e:
            self.response_text.insert(tk.END, f"エラーが発生しました: {e}\n")


root = tk.Tk()
app = App(root)
root.mainloop()