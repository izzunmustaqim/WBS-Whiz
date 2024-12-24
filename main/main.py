import json
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd
import requests  # Import pandas
import pandas as pd  # Import pandas
import config   # Import the config file
import webbrowser
import re
import io
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from threading import Thread  # Import Thread class
import base64  # Import base64 module
from PIL import Image  # Import Image class from PIL module

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid()
        self.skillset_file = None  # Initialize skillset_file
        self.task_details_file = None  # Initialize task_file
        self.create_widgets()
        self.skill_set_data = None  # Initialize skill_set_data
        self.task_details_data = None
        self.ss_folder_file = []
        self.screen_layout_json = None
        self.app_detailed_spec_data_converted_json = None
        self.flowchart_image_path = None

    def create_widgets(self):
        # Add a new entry for the API key
        tk.Label(self, text="API Key").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.api_key_entry = tk.Entry(self, width=40, show='*')
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        
        # Members skillset input
        tk.Label(self, text="Members skill set").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.skillset_entry = tk.Text(self, height=1, width=40)
        self.skillset_entry.grid(row=2, column=1, padx=10, pady=5, sticky='we')
        tk.Button(self, text="Browse", command=lambda e=self.skillset_entry, l="Members skill set": self.browse_file(e, l), width=10).grid(row=2, column=2, padx=10, pady=5)
        
        # Template download link
        text_widget = tk.Text(self, height=1, width=40, font=("Helvetica", 8), bd=0, bg=self.cget("bg"))
        text_widget.grid(row=3, column=1, padx=10, sticky='w')
        text_widget.insert(tk.END, "Download template here: Members_skillset.xlsx")
        text_widget.tag_add("link", "1.23", "1.end")
        text_widget.tag_config("link", foreground="blue", underline=True)
        text_widget.tag_bind("link", "<Button-1>", lambda e, url="https://fujitsu.sharepoint.com/:x:/r/teams/Asia-42f6e454-ChatAIContestAPG/Shared%20Documents/ChatAI%20Contest%20APG/Deliverable/Sprint%201/MEMBERS_SKILLSET.xlsx?d=w3087fb3ba54e43bab309789ad185a9a7&csf=1&web=1&e=5ekEi5": self.open_url(url))

        # Add radio buttons for Task Details and SS
        tk.Label(self, text="Input Details").grid(row=5, column=0, padx=10, pady=5, sticky='w')
        self.file_type = tk.StringVar(value="Task Details")
        tk.Radiobutton(self, text="Task Details", variable=self.file_type, value="Task Details", command=self.update_file_selection).grid(row=5, column=1, padx=2, pady=5, sticky='w')
        tk.Radiobutton(self, text="SS Documents", variable=self.file_type, value="SS Documents", command=self.update_file_selection).grid(row=5, column=1, padx=(100,0), pady=5, sticky='w')
                
         # Text area to display all selected files
        self.input_details_entry = tk.Text(self, height=5, width=40)
        self.input_details_entry.grid(row=6, column=1, padx=10, pady=5)
        tk.Button(self, text="Browse", command=lambda e=self.input_details_entry: self.browse_file(e, self.file_type.get()), width=10).grid(row=6, column=2, padx=10, pady=(0,60))
        
        # Template download link
        text_widget = tk.Text(self, height=1, width=40, font=("Helvetica", 8), bd=0, bg=self.cget("bg"))
        text_widget.grid(row=7, column=1, padx=10, sticky='w')
        text_widget.insert(tk.END, "Download template here: Task_details.xlsx")
        text_widget.tag_add("link", "1.23", "1.end")
        text_widget.tag_config("link", foreground="blue", underline=True)
        text_widget.tag_bind("link", "<Button-1>", lambda e, url="https://fujitsu.sharepoint.com/:x:/r/teams/Asia-42f6e454-ChatAIContestAPG/Shared%20Documents/ChatAI%20Contest%20APG/Deliverable/Sprint%202/Task%20Details%20Sample.xlsx?d=w4ffbd7b446c146539859793651360c36&csf=1&web=1&e=jUauGt": self.open_url(url))

        # Start and End date picker
        tk.Label(self, text="Project Duration").grid(row=9, column=0, padx=10, pady=10, sticky='w')
        #Start date
        tk.Label(self, text="Start:").grid(row=9, column=1, padx=8, pady=10, sticky='w')
        self.start_date_entry = DateEntry(self, width=12, background='darkblue', foreground='white', borderwidth=2, mindate=datetime.now(), date_pattern='MM/dd/yyyy')
        #print(self.start_date_entry.get_date()) --> to get the date value self.start_date_entry.get_date()
        self.start_date_entry.grid(row=9, column=1, padx=(50,0), pady=10, sticky='w')
        #End date
        tk.Label(self, text="End:").grid(row=9, column=1, padx=(160,0), pady=10, sticky='w')
        self.end_date_entry = DateEntry(self, width=12, background='darkblue', foreground='white', borderwidth=2, mindate=datetime.now(), date_pattern='MM/dd/yyyy')
        self.end_date_entry.grid(row=9, column=1, padx=(200,0), pady=10, sticky='w')

        # Bind the validation function to the end date entry
        self.end_date_entry.bind("<<DateEntrySelected>>", self.validate_dates)

        # Start and Cancel button
        tk.Button(self, text="Start", command=self.button_starter, width=10).grid(row=11, column=1, padx=10, pady=10, sticky='e')
        tk.Button(self, text="Cancel", command=self.master.destroy, width=10).grid(row=11, column=2, padx=10, pady=10, sticky='w')
    
    def validate_dates(self, event):
        start_date = self.start_date_entry.get_date()
        end_date = self.end_date_entry.get_date()
        if end_date < start_date:
            messagebox.showerror("Error", "End date cannot be before start date")
            self.end_date_entry.set_date(start_date)

    def update_file_selection(self):
        self.input_details_entry.config(state=tk.NORMAL)
        self.input_details_entry.delete(1.0, tk.END)
        self.input_details_entry.config(state=tk.DISABLED)

    def create_result_section(self):
        # Add a separator
        self.separator = ttk.Separator(self, orient='horizontal')
        self.separator.grid(row=12, column=0, columnspan=3, sticky='we', pady=10)

        # Result section
        self.result_section = tk.Label(self, text="Result: ")
        self.result_section.grid(row=13, column=0, columnspan=3, padx=5, pady=5)

        # Status label to show process completion
        self.status_label = tk.Label(self, text="Processing, please wait...", wraplength=400)
        self.status_label.grid(row=14, column=0, columnspan=3, padx=10, pady=5)
        self.status_label.update_idletasks()  # Force the GUI to update

        # Add a button to save a file and center it
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        self.download_button = tk.Button(self, text="Download WBS", command=self.download_result, width=10)
        self.download_button.grid(row=15, column=1, padx=10, pady=10, sticky='ew')
        self.download_button.grid_remove()

    def remove_result_section(self):
        if hasattr(self, 'separator'):
            self.separator.grid_forget()
        if hasattr(self, 'result_section'):
            self.result_section.grid_forget()
        if hasattr(self, 'status_label'):
            self.status_label.grid_forget()
        if hasattr(self, 'download_button'):
            self.download_button.grid_remove()

    def browse_file(self, entry, label): 
        # if the process is repeated - clear the list first
        self.ss_folder_file.clear()
        selected_file_type = self.file_type.get()
        file_types = [("Excel files", "*.xlsx *.xls")]
        try:
            entry.config(state=tk.NORMAL)
            entry.delete(1.0, tk.END)
            if label == "Members skill set" or selected_file_type == "Task Details":
                file_path = filedialog.askopenfilename(filetypes=file_types)
                if file_path:
                    if label == "Members skill set":
                        self.skillset_file = file_path
                    else:
                        self.task_details_file = file_path

                entry.insert(tk.END, file_path)
                entry.config(state=tk.DISABLED)
            else:
                # Select a folder
                entry.config(state=tk.NORMAL)
                entry.delete(1.0, tk.END)
                folder_path = filedialog.askdirectory()
                
                if folder_path:
                    self.ss_document_folder = folder_path
                    # List all file paths in the selected folder and sub-folders
                    folder_file_paths = []
                    for root, _, files in os.walk(folder_path):
                        for file in files:
                            if file.endswith('.xlsx') or file.endswith('.xls'):
                                folder_file_paths.append(os.path.join(root, file))

                    if len(folder_file_paths) > 5:
                        messagebox.showerror("Error", config.error_message["ManyExcelError"])
                        self.ss_folder_file.clear()  # Clear the list to prevent further processing
                        return
                    else:
                        print("Files in the selected folder and sub-folders:")
                        for file_path in folder_file_paths:
                            entry.insert(tk.END, file_path + "\n")
                            self.ss_folder_file.append(file_path)
                    entry.config(state=tk.DISABLED)
                else:
                    messagebox.showerror("Error", config.error_message["FolderNotFoundError"])
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while selecting the files: {e}")

    def open_url(self, url):
        webbrowser.open_new(url)

    def button_starter(self):
            t = Thread(target=self.main)
            t.start()
    
    def main(self):
        # if repeat the process, clear the result section first
        self.remove_result_section()
        
        self.api_key = self.validate_api_key(self.api_key_entry.get())
        if not self.api_key:
            return False
        
        # Compare excel with template
        # if self.compare_excel(self.skillset_file, 'MEMBERS_SKILLSET.xlsx') == False:
        #     return
        
        # Read skillset data
        self.skill_set_data = self.read_file(self.skillset_file, 3, None)
        if self.skill_set_data is None:
            return  # Stop the process if the file is not found or an error occurred

        # If the file type is Task Details, read the file
        if self.file_type.get() == "Task Details":
            # Read task details data
            if self.compare_excel(self.task_details_file, 'Task Details Sample.xlsx') == False:
                return
            self.task_details_data = self.read_file(self.task_details_file, 1, "B:E")
            if self.task_details_data is None:
                return
        else:
            self.read_ss_folder_files()
            self.request_task_details()

        # print(f"Result in main :")
        # print(self.screen_layout_json)
        # print('---------------------------------------------------------')
        # print(self.app_detailed_spec_data_converted_json)

        # self.send_data_to_chatai()
    
    def compare_excel(self, file, template_file): 
        try:
            workbook1 = load_workbook(file, data_only=True)
            workbook2 = load_workbook(template_file, data_only=True)
            sheet1 = workbook1.active
            sheet2 = workbook2.active

            comparison1_success = True
            comparison2_success = True
            comparison3_success = True
    
            if workbook1 == self.task_details_file:
                # Comparison 1: Column B, rows 2-4
                for row in range(2, 5):  # Rows 2, 3, 4
                    if sheet1[f"B{row}"].value != sheet2[f"B{row}"].value:
                        print(f"Comparison 1 failed: Cell B{row} differs.")
                        comparison1_success = False
    
                # Comparison 2: Columns B-E, row 6
            for col in range(2, 6):  # Columns B, C, D, E (2,3,4,5)
                if sheet1[f"{chr(ord('B') + col - 2)}{6}"].value != sheet2[f"{chr(ord('B') + col - 2)}{6}"].value:
                    print(f"Comparison 2 failed: Cell {chr(ord('B') + col - 2)}{6} differs.")
                    comparison2_success = False
    
            if workbook1 == self.skillset_file:
                # Comparison 3: Columns B-GP, rows 2-4
                for row in range(2, 5):  # Rows 2, 3, 4
                    for col in range(2, 188):  # Columns B (2) to GP (187)
                        col_str = chr(col + 64)  # Convert column number to letter
                        if sheet1.cell(row=row, column=col).value != sheet2.cell(row=row, column=col).value:
                            print(f"Comparison 3 failed: Cell {col_str}{row} differs.")
                            comparison3_success = False

            overall_comparison = comparison1_success and comparison2_success and comparison3_success

            print(f"Overall comparison: {overall_comparison}")
            if overall_comparison != True:
                messagebox.showerror("Overall Comparison", f"Please double check format/input in excel")
                return overall_comparison
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare Excel files: {e}")
            return False
        
    # rangecol=None mean will read all columns that has data
    def read_file(self, file, rowskip=0, rangecol=None):
        # print(file, rowskip, rangecol)
        try:
            file_data = pd.read_excel(file, skiprows=rowskip, usecols=rangecol)  # Read Excel file using pandas
            
            # Check if the file is empty
            if file_data.empty:
                raise pd.errors.EmptyDataError("The file is empty")
            
            # Data cleaning steps
            file_data.dropna(axis=1, how='all', inplace=True)  # Remove columns with all missing values
            file_data.fillna(0, inplace=True)  # Replace NaN values with 0
            
            # print(file_data)  # Print the task details data for debugging
            return(file_data)
        except FileNotFoundError:
            messagebox.showerror("Error", config.error_message["FileNotFoundError"])
            return None
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", config.error_message["EmptyDataError"])
            return None
        except pd.errors.ParserError:
            messagebox.showerror("Error", config.error_message["ParserError"])
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return None

    def read_ss_folder_files(self):
        self.selected_folder_file = []
        
        # Variables to store file paths based on keywords
        application_detailed_specification_files = []
        event_process_diagram_history_files = []
        screen_layout_files = []
        
        for file_path in self.ss_folder_file:
            if "Application Detailed Specification" in file_path:
                application_detailed_specification_files.append(file_path)
            elif "Event Process Sequence Diagram History" in file_path:
                 event_process_diagram_history_files.append(file_path)
            elif "Screen Layout" in file_path:
                screen_layout_files.append(file_path)
        
        # Print the results
        print("\nScreen Layout Files:")
        for file in screen_layout_files:
            sheetName = "項目定義"
            keywordsHeader = ['画面項目名\n/Screen Item Name', 'タイプ\n/ Type']
            self.screen_layout_json = self.read_screen_layout(file, sheetName, keywordsHeader)
            print(f"{self.screen_layout_json}")
            pass

        print("Application Detailed Specification Files:")
        for file in application_detailed_specification_files:
            # print(file)
            # call function read application detailed specification files
            app_detailed_spec_data = self.read_application_detailed_specification_files(file)
            self.app_detailed_spec_data_converted_json = self.convert_app_detailed_spec_data(app_detailed_spec_data)
            print(self.app_detailed_spec_data_converted_json)

        print("\nEvent Process Sequence Diagram History Files:")
        for file in event_process_diagram_history_files:
            # call function read event process sequence diagram history files
            self.flowchart_image_path = self.event_Process_Sequence_Diagram_History_Files(file)
            # print(flowchart_image)
        
    def read_screen_layout(self, file, sheetName, keywordsHeader):
        try:
            workbook = load_workbook(filename=file)
            sheet_names = [] # Initialize an empty list for sheet names

            # Iterate through each worksheet in the workbook
            for sheet in workbook.worksheets:
                sheet_names.append(sheet.title)

            # Check if the file is an Application Detailed Specification file
            self.check_file_validity(sheet_names, workbook, file)

            sheet = workbook[sheetName]  
            screen_layout_data = []
            start_found = False

            for row in sheet.iter_rows(values_only=True):
                filtered_row = []
                for cell in row:
                    if cell is not None:
                        filtered_row.append(cell)

                if any(keyword in str(cell) for keyword in keywordsHeader for cell in filtered_row):
                    start_found = True
                    
                if start_found:
                    if filtered_row != [] and '画面項目名\n/Screen Item Name' not in filtered_row:
                        screen_layout_data.append(filtered_row)

            # Initialize the JSON structure
            screen_layout_json = {
                "Screen Layout": [],
            }
    
            for row in screen_layout_data:
                if len(row) > 1 and (row[1] != '-' or row[2] != '-'):
                    screen_layout_json["Screen Layout"].append({
                        "Screen Item Name": row[1],
                        "Type": row[2],
                    })
            
            # Convert to JSON string for readability
            json_string = json.dumps(screen_layout_json, indent=4)
            return json_string
    
        except FileNotFoundError:
            messagebox.showerror("Error", config.error_message["FileNotFoundError"])
            return None
        except ValueError as e:
            messagebox.showerror("Error", config.error_message["EmptyDataError"])
            return None
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            return None

    def read_application_detailed_specification_files(self, file_path):
        # Read application detailed specification files
        try:
            workbook = load_workbook(filename=file_path, data_only=True) # Load the workbook
            sheet_names = [] # Initialize an empty list for sheet names

            # Iterate through each worksheet in the workbook
            for sheet in workbook.worksheets:
                sheet_names.append(sheet.title)

            # Check if the file is an Application Detailed Specification file
            self.check_file_validity(sheet_names, workbook, file_path)
            
            sheet = workbook[sheet_names[2]] # Select the third sheet
            application_detailed_spec_data = []

            # Define the start and end keywords
            end_keyword = ['メンバ定義\n/Member Definition','メンバ名\n/Member Name']
            start_keywords = ['説明\n/Description', '引数\n/Argument', '戻り値\n/Return Value', 'テーブル/ファイル\n/Table/File']
            start_found = False     # Initialize flags

            # Iterate through rows and print rows between the keywords
            for row in sheet.iter_rows(values_only=True):
                filtered_row = [cell for cell in row if cell is not None]
                # print(filtered_row)
                if any(start_keyword in str(cell) for start_keyword in start_keywords for cell in filtered_row):
                    start_found = True
                if any(end_keyword in str(cell) for end_keyword in end_keyword for cell in filtered_row):
                    start_found = False
                if start_found:
                    if filtered_row != []:
                        # print(filtered_row)
                        application_detailed_spec_data.append(filtered_row)
            return application_detailed_spec_data
        
        except FileNotFoundError:
            messagebox.showerror("Error", config.error_message["FileNotFoundError"])
            return None
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", config.error_message["EmptyDataError"])
            return None
        except pd.errors.ParserError:
            messagebox.showerror("Error", config.error_message["ParserError"])
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return None
    
    def event_Process_Sequence_Diagram_History_Files(self, file_path):
        try:
            sheet_names = [] # Initialize an empty list for sheet names
            workbook = load_workbook(file_path, data_only=True)

            # Iterate through each worksheet in the workbook
            for sheet in workbook.worksheets:
                sheet_names.append(sheet.title)
                print(sheet.title)

            # # Check if the file is an Application Detailed Specification file
            self.check_file_validity(sheet_names, workbook, file_path)
            
            sheet = workbook[sheet_names[2]] # Select the third sheet

            # Extract images from the Excel file
            images = []
            for idx, image in enumerate(sheet._images):
                # Convert the openpyxl image to a PIL image
                image_stream = io.BytesIO(image._data())
                img = Image.open(image_stream)
                images.append(img)
                
                # Save the image to a file
                img.save(f'image_{idx + 1}.png')

            image_path = "image_1.png"
            #description = self.describe_image(image_path)
            # Delete the file
            # os.remove(image_path)
            return image_path
            # print(description)

        except FileNotFoundError:
            messagebox.showerror("Error", config.error_message["FileNotFoundError"])
            return None
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", config.error_message["EmptyDataError"])
            return None
        except pd.errors.ParserError:
            messagebox.showerror("Error", config.error_message["ParserError"])
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return None

    def describe_image(self, img):
        encoded_image = self.encode_image(img)
    
        headers = {
            "Content-type": "application/json",
            "api-key": self.api_key
        }
    
        # data = {
        #     "messages": [
        #         {"role": "user", "content": [
        #             {"type": "text", "text": "Please describe the image below."},
        #             {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{encoded_image}"}}
        #         ]}
        #     ]
        # }

        data = {
            "contents": [
                {
                    "role": "user",
                    "parts": [
                        {
                            "inlineData": {
                                "mimeType": "image/jpeg",
                                "data": encoded_image
                            }
                        },
                        {
                            "text": "What is shown this image?"
                        }
                    ]
                }
            ]
        }
    
        # response = requests.post("https://ai-foundation-api.app/ai-foundation/chat-ai/gpt4", headers=headers, json=data)
        response = requests.post("https://api.ai-service.global.fujitsu.com/ai-foundation/chat-ai/gemini/pro:generateContent", headers=headers, json=data)
        response_json = response.json()
        # print(response_json)
        # return response_json["choices"][0]["message"]["content"]
        return response_json['candidates'][0]['content']['parts'][0]['text']

    def encode_image(self,img):
        with open(img, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    
    def check_file_validity(self, sheet_names, workbook, file_path):
        file_size = os.path.getsize(file_path)
        max_mb = 2000000 * 1024 * 1024  # 2M MB in bytes
        if file_size > max_mb:
            raise ValueError("The file is too large. Please upload a smaller file size")
            # print(f"File size: {file_size} bytes")
        for sheet_name in sheet_names:
                sheet_temp = workbook[sheet_name]
                if sheet_temp.max_row == 1 and sheet_temp.max_column == 1 and sheet_temp.cell(row=1, column=1).value is None:
                    raise ValueError("The file is empty")
                for row in sheet_temp.iter_rows(values_only=True):
                    filtered_row = [cell for cell in row if cell is not None]
                    if filtered_row == []:
                        raise ValueError("The file is empty")
                    elif "Screen Layout" in file_path and '画面レイアウト\n/Screen Layout' not in filtered_row:
                        raise ValueError("The file is not a Screen Layout file")
                    elif "Application Detailed Specification" in file_path and "アプリケーション詳細仕様\n/Application Detailed Specification" not in filtered_row:
                        raise ValueError("The file is not an Application Detailed Specification file")
                    elif "Event Process Sequence Diagram History" in file_path and "イベント処理シーケンス図\n/Event Process Sequence Diagram" not in filtered_row:
                        raise ValueError("The file is not an Event Process Sequence Diagram History file")
                    else:
                        return True
        
    def convert_app_detailed_spec_data(self, app_detailed_spec_data):
        is_description = False
        is_argument = False
        is_return_value = False
        is_table_file = False
        description = []
        argument = []
        return_value = []
        table_file = []
        keywords = ['説明\n/Description', '引数\n/Argument', '戻り値\n/Return Value', 'テーブル/ファイル\n/Table/File']

        for row in app_detailed_spec_data:
            if (keywords[0] in row and '名称\n/Name' not in row) or is_description:
                # Extract the description
                is_table_file = False
                if keywords[1] not in row:
                    is_description = True
                    if len(row) == 2:
                        description.append(row[1])
                    elif len(row) == 1 and row[0] != '説明\n/Description':
                        description.append(row[0])

            if keywords[1] in row or is_argument:
                # Extract the argument
                if keywords[2] not in row:
                    is_description = False
                    is_argument = True
                    if len(row) == 4:
                        if '名称\n/Name' not in row:
                            argument.append(row)

            if keywords[2] in row or is_return_value:
                # Extract the return value
                if keywords[3] not in row:
                    is_argument = False
                    is_return_value = True
                    if len(row) == 4:
                        if '名称\n/Name' not in row:
                            return_value.append(row)
                
            if keywords[3] in row or is_table_file:
                # Extract the table/file
                is_return_value = False
                is_table_file = True
                if len(row) == 7:
                    table_file.append(row)
        
        # Initialize the JSON structure
        app_detailed_spec_json = {
            "Description": [],
            "Argument": [],
            "Return_value": [],
            "Table or File use": []
        }

        for row in description:
            if len(row) > 1:
                app_detailed_spec_json["Description"].append(row)

        for row in argument:
            if len(row) > 1:
                app_detailed_spec_json["Argument"].append({
                    "No": row[0],
                    "Name": row[1],
                    "Type": row[2],
                    "Description": row[3]
                })
        
        for row in return_value:
            if len(row) > 1:
                app_detailed_spec_json["Return_value"].append({
                    "No": row[0],
                    "Name": row[1],
                    "Type": row[2],
                    "Description": row[3]
                })
        
        for row in table_file:
            if len(row) > 1:
                app_detailed_spec_json["Table or File use"].append({
                    "No": row[0],
                    "Table_ID/File_ID": row[1],
                    "Table_Name/File_Name": row[2],
                    "CRUD Access for C": row[3],
                    "CRUD Access for R": row[4],
                    "CRUD Access for U": row[5],
                    "CRUD Access for D": row[6]
                })
        
        # Convert to JSON string for readability
        json_string = json.dumps(app_detailed_spec_json, indent=4, ensure_ascii=False)
        # print(json_string)
        return json_string

    def request_task_details(self):
        try:
            # Call the method to create the status section
            self.create_result_section()

            encoded_image = self.encode_image(self.flowchart_image_path)
    
            headers = {
                "Content-type": "application/json",
                "api-key": self.api_key
            }
            prompt = config.prompt_list_task.format(
                            screen_layout_json=self.screen_layout_json,
                            app_detailed_spec_data_converted_json=self.app_detailed_spec_data_converted_json
                        )
            api_endpoint = "https://api.ai-service.global.fujitsu.com/ai-foundation/chat-ai/gemini/pro:generateContent" 

            data = {
                "contents": [
                    {
                        "role": "user",
                        "parts": [
                            {
                                "inlineData": {
                                    "mimeType": "image/jpeg",
                                    "data": encoded_image
                                }
                            },
                            {
                                "text": prompt
                            }
                        ]
                    }
                ]
            }
        
            # Send the POST request
            response = requests.post(api_endpoint, headers=headers, json=data)
            response.raise_for_status()  # Raise an exception for HTTP errors
            os.remove(self.flowchart_image_path)
           
            # Check the response
            try:
                analysis_result = response.json()
                # print("Analysis Result:", analysis_result)
                # Extract the content (only the wbs result)
                content = analysis_result['candidates'][0]['content']['parts'][0]['text']
                # self.create_wbs(content, start_date)
                print("Response from chat AI for the Task Details")
                print(content)
 
            except json.JSONDecodeError:
                print("Error: The response is not in JSON format.")
                print("Response content:", response.text)
 
        except requests.exceptions.RequestException as e:
            if "Too Large" in str(e):
                messagebox.showerror("Error", config.error_message["FileTooBig"])
            else:
                print("Failed to get a response from ChatAI. Status code:", response.status_code)
                print("Response content:", response.text)
                messagebox.showerror("Error", f"Failed to send data to ChatAI: {e}")

        except ValueError as ve:
            print(ve)
            messagebox.showerror("Error", str(ve))
        finally:
            self.status_label.config(text="Process Task Details has completed successfully. Creating WBS is in progress.")
    
    #Send data to ChatAI for analysis
    def send_data_to_chatai(self):
        try:
 
            # Call the method to create the status section
            self.create_result_section()
           
            # Load the Excel file
            workbook = load_workbook(self.task_details_file)
 
            # Select the active sheet (or specify a sheet name)
            sheet = workbook.active
 
            # Read the value from column C, row 2
            start_date = sheet['C2'].value
            end_date = sheet['C3'].value
 
            # Debug - print the start date and end date
            print(start_date)
            print(end_date)
 
            # Define the API endpoint and hardcoded prompt
            api_endpoint = "https://api.ai-service.global.fujitsu.com/ai-foundation/chat-ai/gemini/pro:generateContent" 
            prompt = config.prompt.format(
                            task_details_data=self.task_details_data.to_json(),
                            skill_set_data=self.skill_set_data.to_json(),
                            start_date_str=start_date,
                            end_date_str=end_date,
                            task_description="Task Description Example",  # Provide example values for placeholders
                            assigned_to="Assigned to Example",
                            progress="To do",
                            plan_start_date="Start date Example",
                            plan_end_date="End date Example"
                        )
 
            headers = {
                "Content-type": "application/json",
                "api-key": self.api_key
            }
 
            payload = {
                "contents": [
                {
                    "role": "user",
                    "parts": [
                        {
                            "text": prompt
                        }
                    ]
                }
                ]
            }
           
            # Send the POST request
            response = requests.post(api_endpoint, headers=headers, json=payload)
            response.raise_for_status()  # Raise an exception for HTTP errors
           
            # Check the response
            try:
                analysis_result = response.json()
                #print("Analysis Result:", analysis_result)
 
                # Extract the content (only the wbs result)
                content = analysis_result['candidates'][0]['content']['parts'][0]['text']
                self.create_wbs(content, start_date)
                print(content)
 
            except json.JSONDecodeError:
                print("Error: The response is not in JSON format.")
                print("Response content:", response.text)
 
        except requests.exceptions.RequestException as e:
            if "Too Large" in str(e):
                messagebox.showerror("Error", config.error_message["FileTooBig"])
            else:
                print("Failed to get a response from ChatAI. Status code:", response.status_code)
                print("Response content:", response.text)
                messagebox.showerror("Error", f"Failed to send data to ChatAI: {e}")

        except ValueError as ve:
            print(ve)
            messagebox.showerror("Error", str(ve))
        finally:
            self.status_label.config(text="Process has completed successfully. You may download the WBS file using the download button below.")
            self.download_button.grid()

    def download_result(self):
        try:
            # Define the destination file path in the Downloads folder
            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            destination_file_path = os.path.join(downloads_folder, "Details_WBS.xlsx")

            # Create dummy file in the download folder
            df = pd.DataFrame()
            df.to_excel(destination_file_path, index=False)

            # Get the current directory
            current_directory = os.getcwd()

            # Define the source file path
            source_file_path = os.path.join(current_directory, "Details_WBS.xlsx")

            # Copy the file
            shutil.copy(source_file_path, destination_file_path)
            messagebox.showinfo("Success", f"File saved successfully to {destination_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

    def validate_api_key(self, api_key):
        pattern = r'^[A-Za-z0-9]{48}$'        
        if not api_key:
            messagebox.showerror("Error", config.error_message["APIEmptyField"])
            return None
        elif re.search(r'[\uFF01-\uFF60\uFFE0-\uFFE6]', api_key):
            messagebox.showerror("Error", config.error_message["FullWidthCharacterError"])
            return None
        elif re.match(pattern, api_key):
            return api_key
        else:
            messagebox.showerror("Error", config.error_message["InvalidKeyError"])
            return None
        
    def create_wbs(self, content, start_date):
        # Extract the markdown table using regular expression
        table_pattern = re.compile(r'\|.*\|')
        markdown_table = '\n'.join(table_pattern.findall(content))

        # Convert the Markdown table to a DataFrame
        data = io.StringIO(markdown_table)
        df = pd.read_csv(data, sep="|", skipinitialspace=True, engine='python')

        # Remove the first row
        df = df.iloc[1:]

        # Drop the first and last columns which are empty due to the table format
        df = df.drop(df.columns[[0, -1]], axis=1)

        # Convert to datetime and change format
        df.iloc[:, 4] = pd.to_datetime(df.iloc[:, 4]).dt.strftime('%m/%d/%Y')
        df.iloc[:, 5] = pd.to_datetime(df.iloc[:, 5]).dt.strftime('%m/%d/%Y')

        # remove the header from an existing Pandas DataFrame
        df = df.rename(columns=df.iloc[0]).drop(df.index[0])
        print(df)

        # Load the Excel template
        template_path = 'JDU-WBS_Template_Samples.xlsx'
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active  # or specify the sheet name with wb['SheetName']

        # Write the variable into cell
        ws['B2'] = "Details_WBS.xlsx"
        ws['B6'] = datetime.strptime(start_date, "%Y-%m-%d").strftime("%m/%d/%Y")
        # Set current_date to the current date
        current_date = datetime.now().date()
        ws['G2'] = current_date

         # Write the DataFrame to the Excel template starting at row 9
        start_row = 9
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Save the modified template as a new file
        output_path = 'Details_WBS.xlsx'
        wb.save(output_path)

        print(f"DataFrame saved to {output_path}")


root = tk.Tk()
root.title("WBS Enhancement")
app = Application(master=root)
app.mainloop()