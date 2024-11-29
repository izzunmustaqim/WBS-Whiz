import json
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook
import pandas as pd
import requests  # Import pandas
import pandas as pd  # Import pandas
import config   # Import the config file
import webbrowser
import re
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from threading import Thread  # Import Thread class

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid()
        self.skillset_file = None  # Initialize skillset_file
        self.task_details_file = None  # Initialize task_file
        self.create_widgets()
        self.skill_set_data = None  # Initialize skill_set_data

    def create_widgets(self):
        # Add a new entry for the API key
        tk.Label(self, text="API Key").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.api_key_entry = tk.Entry(self, width=40, show='*')
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')

        labels = ["Members skill set", "Task details"]
        descriptions = [
            ("Download template here: Members_skillset.xlsx", "https://fujitsu.sharepoint.com/:x:/r/teams/Asia-42f6e454-ChatAIContestAPG/Shared%20Documents/ChatAI%20Contest%20APG/Deliverable/Sprint%201/MEMBERS_SKILLSET.xlsx?d=w3087fb3ba54e43bab309789ad185a9a7&csf=1&web=1&e=5ekEi5"),
            ("Download template here: Task_details.xlsx", "https://fujitsu.sharepoint.com/:x:/r/teams/Asia-42f6e454-ChatAIContestAPG/Shared%20Documents/ChatAI%20Contest%20APG/Deliverable/Sprint%202/Task%20Details%20Sample.xlsx?d=w4ffbd7b446c146539859793651360c36&csf=1&web=1&e=jUauGt")
        ]
        self.entries = []
        for i, (label, (description, url)) in enumerate(zip(labels, descriptions)):
            tk.Label(self, text=label).grid(row=(i+1)*2, column=0, padx=10, pady=5, sticky='w')
            entry = tk.Entry(self, width=40)
            entry.grid(row=(i+1)*2, column=1, padx=10, pady=5)
            self.entries.append(entry)
            tk.Button(self, text="Browse", command=lambda e=entry, l=label: self.browse_file(e, l), width=10).grid(row=(i+1)*2, column=2, padx=10, pady=5)
            
            text_widget = tk.Text(self, height=1, width=40, font=("Helvetica", 8), bd=0, bg=self.cget("bg"))
            text_widget.grid(row=(i+1)*2+1, column=1, padx=10, sticky='w')
            text_widget.insert(tk.END, description)
            
            if "Members_skillset.xlsx" in description:
                text_widget.tag_add("link", "1.23", "1.end")
                text_widget.tag_config("link", foreground="blue", underline=True)
                text_widget.tag_bind("link", "<Button-1>", lambda e, url=url: self.open_url(url))
            else:
                text_widget.tag_add("link", "1.23", "1.end")
                text_widget.tag_config("link", foreground="blue", underline=True)
                text_widget.tag_bind("link", "<Button-1>", lambda e, url=url: self.open_url(url))
            
            text_widget.config(state=tk.DISABLED)

            tk.Button(self, text="Start", command=self.button_starter, width=10).grid(row=8, column=1, padx=10, pady=10, sticky='e')
            tk.Button(self, text="Cancel", command=self.master.destroy, width=10).grid(row=8, column=2, padx=10, pady=10, sticky='w')
    
    def create_result_section(self):
        # Add a separator
        self.separator = ttk.Separator(self, orient='horizontal')
        self.separator.grid(row=9, column=0, columnspan=3, sticky='we', pady=10)

        # Result section
        self.result_section = tk.Label(self, text="Result: ")
        self.result_section.grid(row=10, column=0, columnspan=3, padx=5, pady=5)

        # Status label to show process completion
        self.status_label = tk.Label(self, text="Processing, please wait...", wraplength=400)
        self.status_label.grid(row=11, column=0, columnspan=3, padx=10, pady=5)
        self.status_label.update_idletasks()  # Force the GUI to update

        # Add a button to save a file and center it
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        tk.Button(self, text="Download WBS", command=self.download_result, width=10).grid(row=12, column=1, padx=10, pady=10, sticky='ew')

    def remove_result_section(self):
        if hasattr(self, 'separator'):
            self.separator.grid_forget()
        if hasattr(self, 'result_section'):
            self.result_section.grid_forget()
        if hasattr(self, 'status_label'):
            self.status_label.grid_forget()

    def browse_file(self, entry, label):
        file_types = [("Excel files", "*.xlsx *.xls")]
        file_path = filedialog.askopenfilename(filetypes=file_types)
        
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)
            
            if label == "Members skill set":
                self.skillset_file = file_path
                print(self.skillset_file)
            else:
                self.task_details_file = file_path
                print(self.task_details_file)
        else:
            messagebox.showerror("Error", "No file selected")

    def open_url(self, url):
        webbrowser.open_new(url)

    def button_starter(self):
            t = Thread(target=self.main)
            t.start()
    
    def main(self):
        self.api_key = self.validate_api_key(self.api_key_entry.get())
        if not self.api_key:
            return

        # Call read file function
        self.skill_set_data = self.read_file(self.skillset_file, 3, None)
        if self.skill_set_data is None:
            return  # Stop the process if the file is not found or an error occurred

        self.task_details_data = self.read_file(self.task_details_file, 1, "B:E")
        if self.task_details_data is None:
            return 
        
        self.send_data_to_chatai()

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
            
            print(file_data)  # Print the task details data for debugging
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

    #Send data to ChatAI for analysis
    def send_data_to_chatai(self):
        try:
            # if repeat the process, clear the result section first
            self.remove_result_section()
 
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
            print("Failed to get a response from ChatAI. Status code:", response.status_code)
            print("Response content:", response.text)
            messagebox.showerror("Error", f"Failed to send data to ChatAI: {e}")

        except ValueError as ve:
            print(ve)
            messagebox.showerror("Error", str(ve))
        finally:
            self.status_label.config(text="Process has completed successfully. You may download the WBS file using the download button below.")
  
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