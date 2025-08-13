import os
import base64
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import threading
import json
import tempfile
import atexit
import subprocess
import shutil


# ---------------- Utility Functions ---------------- #

def encode_image(image_path):
    """Convert image to base64 string."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')


def analyze_bill_with_gpt4_vision(image_path, output_text_widget, api_key=None):
    """Send image to GPT-4o Vision API and parse JSON response."""
    if not api_key:
        output_text_widget.insert(tk.END, "Error: No API key provided.\n")
        output_text_widget.see(tk.END)
        return None

    base64_image = encode_image(image_path)
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": """
                            Extract the following details from this utility bill image:
                            Customer Name, Account Number, Due Date, and Total Amount Due.
                            If any information is not found, state 'N/A'.
                            Provide ONLY JSON in this format:
                            {"Customer Name": "[Name]", "Account Number": "[Number]",
                            "Due Date": "[Date]", "Total Amount Due": "[Amount]", "Payable Within Due Date": "[Amount]", "Payable After Due Date": "[Amount]"}.
                        """
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 300
    }

    try:
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload
        )
        response.raise_for_status()
        response_data = response.json()

        # Extract JSON from model's text
        content = response_data["choices"][0]["message"]["content"]
        json_start = content.find("{")
        json_end = content.rfind("}") + 1
        if json_start != -1 and json_end != -1:
            try:
                json_string = content[json_start:json_end]
                parsed_data = json.loads(json_string)
                # Ensure all expected keys are present
                expected_keys = ["Customer Name", "Account Number", "Due Date", "Total Amount Due", "Payable Within Due Date", "Payable After Due Date"]
                for key in expected_keys:
                    if key not in parsed_data:
                        parsed_data[key] = "N/A"
                return parsed_data
            except json.JSONDecodeError as e:
                output_text_widget.insert(tk.END, f"JSON decode error for {image_path}: {e}\n")
                output_text_widget.insert(tk.END, f"Raw content: {content}\n")
                output_text_widget.see(tk.END)
                return None
        else:
            output_text_widget.insert(tk.END, f"No JSON found in GPT response for {image_path}.\n")
            output_text_widget.insert(tk.END, f"Raw content: {content}\n")
            output_text_widget.see(tk.END)
            return None

    except requests.exceptions.HTTPError as http_err:
        output_text_widget.insert(tk.END, f"HTTP error for {image_path}: {http_err}\n")
        output_text_widget.insert(tk.END, f"Response: {response.text}\n")
        output_text_widget.see(tk.END)
        return None
    except requests.exceptions.RequestException as e:
        output_text_widget.insert(tk.END, f"Request error for {image_path}: {e}\n")
        output_text_widget.see(tk.END)
        return None


def create_excel_sheet(data, output_text_widget=None):
    """Create Excel with all extracted bills and total payable amount."""


    if not data:
        if output_text_widget:
            output_text_widget.insert(tk.END, "No data to write to Excel.\n")
            output_text_widget.see(tk.END)
        return


    columns = ["Filename", "Customer Name", "Account Number", "Due Date", "Total Amount Due", "Payable Within Due Date", "Payable After Due Date"]
    df = pd.DataFrame(data, columns=columns)

    # Ensure Total Amount Due, Payable Within Due Date, and Payable After Due Date are numeric
    df["Total Amount Due"] = pd.to_numeric(df["Total Amount Due"], errors="coerce").fillna(0)
    df["Payable Within Due Date"] = pd.to_numeric(df["Payable Within Due Date"], errors="coerce").fillna(0)
    df["Payable After Due Date"] = pd.to_numeric(df["Payable After Due Date"], errors="coerce").fillna(0)

    # Calculate totals
    total_payable_amount = df["Total Amount Due"].sum()
    total_payable_within_due_date = df["Payable Within Due Date"].sum()
    total_payable_after_due_date = df["Payable After Due Date"].sum()

    wb = Workbook()
    ws = wb.active
    ws.title = "Bill Details"

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    # Adjust column widths
    for column in ws.columns:
        max_length = 0
        column_name = column[0].value # Get the column header
        if column_name is not None:
            max_length = len(str(column_name))
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

    ws.append([])
    ws.append(["Grand Total Payable Amount:", total_payable_amount])
    ws.append(["Total Payable Within Due Date:", total_payable_within_due_date])
    ws.append(["Total Payable After Due Date:", total_payable_after_due_date])

    try:
        # Create a temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_file_path = temp_file.name
        temp_file.close()

        wb.save(temp_file_path)
        if output_text_widget:
            output_text_widget.insert(tk.END, f"Excel file created temporarily at: {temp_file_path}\n")
            output_text_widget.see(tk.END)
        return temp_file_path
    except Exception as e:
        if output_text_widget:
            output_text_widget.insert(tk.END, f"Error saving Excel: {e}\n")
            output_text_widget.see(tk.END)
        return None


# ---------------- Tkinter GUI App ---------------- #

class BillProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Utility Bill Processor")

        self.image_files = []
        self.api_key = tk.StringVar()

        # API Key Frame
        api_frame = tk.Frame(root)
        api_frame.pack(pady=5)
        tk.Label(api_frame, text="OpenAI API Key:").pack(side=tk.LEFT)
        tk.Entry(api_frame, textvariable=self.api_key, width=50, show="*").pack(side=tk.LEFT, padx=5)
        tk.Button(api_frame, text="Save API Key", command=self.save_api_key).pack(side=tk.LEFT, padx=5)

        # File Selection
        file_frame = tk.Frame(root)
        file_frame.pack(pady=5)
        tk.Button(file_frame, text="Select Bill Images", command=self.browse_files).pack(side=tk.LEFT)
        tk.Label(file_frame, text="Selected files:").pack(side=tk.LEFT, padx=5)
        self.file_count_label = tk.Label(file_frame, text="0 files selected")
        self.file_count_label.pack(side=tk.LEFT)

        # Process Button
        tk.Button(root, text="Process Bills", command=self.start_processing_thread).pack(pady=10)

        # Output Area
        self.output_text = scrolledtext.ScrolledText(root, width=80, height=20, wrap=tk.WORD)
        self.output_text.pack(pady=10)

        self.load_api_key()
        self.temp_excel_files = []
        atexit.register(self._cleanup_temp_files)

    def save_api_key(self):
        key = self.api_key.get().strip()
        if key:
            with open(".env", "w") as f:
                f.write(f"OPENAI_API_KEY={key}")
            messagebox.showinfo("Success", "API Key saved.")
        else:
            messagebox.showwarning("Warning", "No API key entered.")

    def load_api_key(self):
        if os.path.exists(".env"):
            with open(".env", "r") as f:
                line = f.readline().strip()
                if line.startswith("OPENAI_API_KEY="):
                    self.api_key.set(line.split("=", 1)[1])
                    messagebox.showinfo("Loaded", "API Key loaded from file.")

    def browse_files(self):
        filetypes = (
            ('Image files', '*.png *.jpg *.jpeg *.gif *.bmp *.tiff'),
            ('All files', '*.*')
        )
        files = filedialog.askopenfilenames(title="Select bill images", filetypes=filetypes)
        if files:
            self.image_files = list(files)
            self.file_count_label.config(text=f"{len(self.image_files)} files selected")
            self.output_text.insert(tk.END, f"Selected {len(self.image_files)} files:\n")
            for f in self.image_files:
                self.output_text.insert(tk.END, f"- {f}\n")
            self.output_text.see(tk.END)

    def start_processing_thread(self):
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, "Starting bill processing...\n")
        self.output_text.see(tk.END)
        threading.Thread(target=self.process_bills).start()

    def process_bills(self):
        if not self.image_files:
            messagebox.showwarning("Input Error", "No bill images selected.")
            return

        api_key = self.api_key.get().strip()
        if not api_key:
            messagebox.showwarning("Input Error", "Please enter your API key.")
            return

        self.output_text.insert(tk.END, f"Processing {len(self.image_files)} files...\n")
        bills_data = []
        for img_path in self.image_files:
            self.output_text.insert(tk.END, f"Analyzing {os.path.basename(img_path)}...\n")
            data = analyze_bill_with_gpt4_vision(img_path, self.output_text, api_key)
            if data:
                data["Filename"] = os.path.basename(img_path)
                bills_data.append(data)
            else:
                self.output_text.insert(tk.END, f"Failed to extract from {os.path.basename(img_path)}\n")
            self.output_text.see(tk.END)

        if bills_data:
            self.output_text.insert(tk.END, "Creating Excel...\n")
            excel_file_path = create_excel_sheet(bills_data, output_text_widget=self.output_text)
            if excel_file_path:
                self.temp_excel_files.append(excel_file_path)
                self.output_text.insert(tk.END, f"\nExcel file available: ")
                self.output_text.insert(tk.END, "Click to Open", "link")
                self.output_text.tag_config("link", foreground="blue", underline=1)
                self.output_text.tag_bind("link", "<Button-1>", lambda e, path=excel_file_path: self._open_excel_file(path))
                self.output_text.insert(tk.END, f"\n")

                # Add a Save As button
                save_as_button = tk.Button(self.root, text="Save Excel As...", command=lambda: self._save_excel_as(excel_file_path))
                self.output_text.window_create(tk.END, window=save_as_button)
                self.output_text.insert(tk.END, "\n")

            messagebox.showinfo("Complete", "Bill processing finished.")
        else:
            self.output_text.insert(tk.END, "No bill data extracted.\n")
            messagebox.showinfo("Complete", "No data extracted.")


    def _open_excel_file(self, file_path):
        try:
            os.startfile(file_path) # For Windows
        except AttributeError:
            subprocess.call(['open', file_path]) # For macOS
        except FileNotFoundError:
            messagebox.showerror("Error", f"File not found: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")

    def _save_excel_as(self, source_path):
        if not os.path.exists(source_path):
            messagebox.showerror("Error", "No Excel file to save. Please process bills first.")
            return

        initial_filename = os.path.basename(source_path)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=initial_filename
        )
        if file_path:
            try:
                import shutil
                shutil.copy(source_path, file_path)
                messagebox.showinfo("Success", f"Excel file saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save Excel file: {e}")

    def _cleanup_temp_files(self):
        for f_path in self.temp_excel_files:
            try:
                if os.path.exists(f_path):
                    os.remove(f_path)
                    print(f"Cleaned up temporary file: {f_path}")
            except Exception as e:
                print(f"Error cleaning up {f_path}: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = BillProcessorApp(root)
    root.mainloop()
