import tkinter as tk
from tkinter import filedialog, messagebox
# from tkinterdnd2 import TkinterDnD, DND_FILES
import pandas as pd
import re
import datetime
from datetime import datetime as dt, time as dt_time
from collections import defaultdict
from fpdf import FPDF

class FileUploadApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Compare")
        self.root.geometry("600x400")
        # self.root.configure(bg="#e0f7fa")
        self.root.configure(bg="#cacaca")

        # Frame with border for the entire screen
        self.main_frame = tk.Frame(root, bg="#ffffff", bd=5, relief="solid")
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        self.title_label = tk.Label(self.main_frame, text="Attendance Checker", font=("Helvetica", 16), bg="#e0f7fa")
        self.title_label.pack(pady=10)
        self.title_label = tk.Label(self.main_frame, text="Please choose the file carefully!", font=("Helvetica", 8), bg="#ff3434")
        self.title_label.pack(pady=6)

        # Manual File Upload
        self.manual_file_frame = tk.Frame(self.main_frame, bg="#e0f7fa", bd=2, relief="solid")
        self.manual_file_frame.pack(pady=5, padx=10, fill="x")

        self.manual_file_label = tk.Label(self.manual_file_frame, text="Upload Manual Sheet:", font=("Helvetica", 12), bg="#e0f7fa")
        self.manual_file_label.pack(side="left", padx=5)

        self.manual_file_path = tk.StringVar()
    
        self.manual_file_button = tk.Button(self.manual_file_frame, text="Choose File", command=self.load_manual_file, bg="#4CAF50", fg="white", font=("Helvetica", 12), bd=0, padx=10, pady=5)
        self.manual_file_button.pack(side="right", padx=5)

        # Machine File Upload
        self.machine_file_frame = tk.Frame(self.main_frame, bg="#e0f7fa", bd=2, relief="solid")
        self.machine_file_frame.pack(pady=5, padx=10, fill="x")

        self.machine_file_label = tk.Label(self.machine_file_frame, text="Upload Machine Sheet:", font=("Helvetica", 12), bg="#e0f7fa")
        self.machine_file_label.pack(side="left", padx=5)

        self.machine_file_path = tk.StringVar()
    
        self.machine_file_button = tk.Button(self.machine_file_frame, text="Choose File", command=self.load_machine_file, bg="#4CAF50", fg="white", font=("Helvetica", 12), bd=0, padx=10, pady=5)
        self.machine_file_button.pack(side="right", padx=5)

        # Process Data Button
        self.process_button = tk.Button(self.main_frame, text="Process Data", command=self.process_data, bg="#4CAF50", fg="white", font=("Helvetica", 12), bd=0, padx=10, pady=5)
        self.process_button.pack(pady=20)

        # Uploading Section with Drag-and-Drop
        self.uploading_frame = tk.Frame(self.main_frame, bg="#e0f7fa", bd=0.5, relief="solid")
        self.uploading_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.upload_list = tk.Listbox(self.uploading_frame, bg="#ffffff", fg="#000000", font=("Helvetica", 12), bd=0)
        self.upload_list.pack(fill="both", expand=True, padx=10, pady=5)
  

    def load_manual_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.manual_file_path.set(file_path)
            self.upload_list.insert(tk.END, f"Manual File: {file_path}")

    def load_machine_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.machine_file_path.set(file_path)
            self.upload_list.insert(tk.END, f"Machine File: {file_path}")

    def extract_numbers_from_string(self, string):
        return re.findall(r'\d+', string)

    def process_data(self):
        if not self.manual_file_path.get() or not self.machine_file_path.get():
            messagebox.showerror("Error", "Please select both manual and machine files.")
            return

        try:
            # Load the data
            machine_df = pd.read_excel(self.machine_file_path.get())
            manual_df = pd.read_excel(self.manual_file_path.get())

            emp_id = manual_df.iloc[2][3]
            manual_df_skip = pd.read_excel(self.manual_file_path.get(), skiprows=20)
            machine_df_skip = pd.read_excel(self.machine_file_path.get(), skiprows=1)

            manualSheetEmployye = {
                "employee": {
                    "id": "",
                    "montly_attendence": []
                }
            }

            manualSheetEmployye["employee"]["id"] = self.extract_numbers_from_string(emp_id)[0]

            for index, row in manual_df_skip.iterrows():
                day = row["Day"]
                time_in = row['Time In']
                time_out = row['Time Out']
                date = row['Date']

                if pd.isna(time_in) and pd.isna(time_out) and pd.isna(date) and pd.isna(day):
                    continue

                if day == "Day" and date == "Date" and time_in == "Time In" and time_out == "Time Out":
                    continue

                manualSheetEmployye["employee"]["montly_attendence"].append({
                    "day": day,
                    "date": date.strftime("%m/%d/%Y") if isinstance(date, dt) else None,
                    "check_in": time_in.strftime("%H:%M") if isinstance(time_in, dt_time) else None,
                    "check_out": time_out.strftime("%H:%M") if isinstance(time_out, dt_time) else None,
                })

            machine_dataset = []
            emp_id = 0
            employee_dict = {
                'id': 0,
                'monthly_attendance': []
            }
            previous_day = None
            previous_checkout_status = None

            for index, row in machine_df.iterrows():
                if emp_id != row["Emp ID"]:
                    grouped_data = defaultdict(lambda: {"day": None, "date": None, "check_in": None, "check_out": None})

                    for item in employee_dict['monthly_attendance']:
                        key = (item["day"], item["date"])
                        if not grouped_data[key]["day"]:
                            grouped_data[key]["day"] = item["day"]
                            grouped_data[key]["date"] = item["date"]
                        grouped_data[key]["check_in"] = grouped_data[key]["check_in"] or item["check_in"]
                        grouped_data[key]["check_out"] = grouped_data[key]["check_out"] or item["check_out"]

                    merged_data = list(grouped_data.values())

                    employee_dict['monthly_attendance'] = merged_data

                    emp_id = row["Emp ID"]

                    machine_dataset.append(employee_dict)
                    employee_dict = {
                        'id': emp_id,
                        'monthly_attendance': []
                    }

                datetime_object = dt.strptime(row["Time"], "%m/%d/%Y %H:%M:%S")

                day_of_week = datetime_object.strftime("%A")
                time_str = datetime_object.strftime("%H:%M")
                date_str = datetime_object.strftime("%m/%d/%Y")
                inout_status = row["Attendance State"]

                if previous_day is not None and previous_day == day_of_week and previous_checkout_status == inout_status:
                    previous_day = None
                    previous_checkout_status = None
                    continue

                previous_day = day_of_week
                previous_checkout_status = 1 if inout_status == 1 else None

                employee_dict['monthly_attendance'].append({
                    "day": day_of_week,
                    "date": date_str,
                    "check_in": time_str if inout_status == 0 else None,
                    "check_out": time_str if inout_status == 1 else None
                })

            if len(machine_dataset) > 0:
                machine_dataset.pop(0)

            machine_value = next((item for item in machine_dataset if item["id"] == int(manualSheetEmployye["employee"]["id"])), None)
            manual_value = manualSheetEmployye["employee"]

            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not save_path:
                return

            pdf = FPDF(orientation="L",format="A3")
            pdf.add_page()
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(50, 15, txt="Attendance Comparison Report", ln=True, align='Center' )
            pdf.set_font("Arial", size=12)
            pdf.cell(50, 10, txt="Red= data mismatched , Green= matched in both sheets, Orange=attendance is missing in one sheet , Black= no data in both sheets", ln=True, align='b')

            headers = [
                'Day', 'Date', 'Manual Check In', 'Manual Check Out',
                'Machine Check In', 'Machine Check Out', 'Check In Status', 'Check Out Status'
            ]
            pdf.set_font("Arial", 'B', 12)
            for header in headers:
                pdf.cell(50, 10, header, border=1, ln=0, align='C')
            pdf.ln()

            pdf.set_font("Arial", size=12)
            for manual_item in manual_value["montly_attendence"]:
                search = next((machine_item for machine_item in machine_value['monthly_attendance'] if manual_item['date'] == machine_item['date']), None)

                data = {
                    'day': manual_item['day'],
                    'date': manual_item['date'],
                    'manual_check_in': manual_item['check_in'],
                    'manual_check_out': manual_item['check_out'],
                    'machine_check_in': search['check_in'] if search is not None else None,
                    'machine_check_out': search['check_out'] if search is not None else None
                }

                if search is None:
                    data['status'] = 'missing in both sheets'
                    pdf.set_text_color(0, 0, 0)

                else:
                    # Check in
                    if search['check_in'] is None:
                        data['check_in_status'] = 'missing in machine sheet'
                    
                        
                    if manual_item['check_in'] is None:
                        data['check_in_status'] = 'Not entered in one sheet'
                    elif manual_item['check_in'] == search['check_in']:
                        data['check_in_status'] = 'Check in matched'
                       
                        
                    else:
                        data['check_in_status'] = 'Check in mismatched'
                        
                        	
                    # Check out
                    if search['check_out'] is None:
                        data['check_out_status'] = 'missing in machine sheet'
                    elif manual_item['check_out'] is None:
                        data['check_out_status'] = 'Data not entered'
                    elif manual_item['check_out'] == search['check_out']:
                        data['check_out_status'] = 'Check out matched'
                    else:
                        data['check_out_status'] = 'Check out mismatched'
                        
                        
                    #set color
                    if data['check_in_status'] == 'Check in matched' and data['check_out_status'] == 'Check out matched':
                        pdf.set_text_color(0, 128, 0)
                    if data['check_in_status'] == 'check in matched' or data['check_out_status'] == 'missing in machine sheet':
                        pdf.set_text_color(255, 100, 50)
                    
                    if data['check_in_status'] == 'Check in mismatched' or data['check_out_status'] == 'Check out mismatched':
                        pdf.set_text_color(255,0,0)
                    if data['check_in_status'] == 'Data not entered':
                        pdf.set_text_color(0, 0, 255)

                # # pdf.set_fill_color(pdf.cell)    

                for value in data.values():
                    pdf.cell(50, 10, str(value), border=1, ln=0, align='C')
                pdf.ln()

            pdf.output(save_path)

            messagebox.showinfo("Success", f"Data processed and saved to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()  # for dragand drop Use TkinterDnD.Tk instead of tk.Tk
    app = FileUploadApp(root)
    root.mainloop()
