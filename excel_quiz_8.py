import customtkinter as ctk
import tkinter.messagebox as msg
import sqlite3
from datetime import datetime
import time
from openpyxl import Workbook
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import styles
from reportlab.lib.units import inch

from reportlab.platypus import Image
import os

from reportlab.platypus import KeepTogether
from reportlab.lib.pagesizes import A4

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "1234"


class ExcelQuizApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # -------------------------------------------------
        # Window Setup
        # -------------------------------------------------
        self.title("Excel Professional Timed Quiz System - By Tryfon Papadopoulos")
        self.center_window(1000, 750)

        # -------------------------------------------------
        # Database Initialization
        # -------------------------------------------------
        self.init_database()

        # -------------------------------------------------
        # Questions
        # -------------------------------------------------
        self.questions = [
            {"question": "What is Excel mainly used for?",
             "options": ["Video editing", "Spreadsheet calculations", "Gaming", "Drawing"],
             "answer": 1},

            {"question": "Formulas start with?",
             "options": ["#", "@", "=", "$"],
             "answer": 2},

            {"question": "Function to add numbers?",
             "options": ["ADD()", "SUM()", "PLUS()", "TOTAL()"],
             "answer": 1},

            {"question": "Intersection of row and column?",
             "options": ["Table", "Sheet", "Cell", "Range"],
             "answer": 2},

            {"question": "Function for average?",
             "options": ["AVG()", "AVERAGE()", "MEAN()", "MID()"],
             "answer": 1},

            {"question": "CTRL + C does?",
             "options": ["Copy", "Close", "Chart", "Calculate"],
             "answer": 0},

            {"question": "Sort is used to?",
             "options": ["Format", "Organize data", "Delete data", "Print"],
             "answer": 1},

            {"question": "Fill Handle is used to?",
             "options": ["Delete", "Extend patterns", "Chart", "Lock cells"],
             "answer": 1},

            {"question": "Charts are created from?",
             "options": ["Insert Tab", "Review", "Data", "View"],
             "answer": 0},

            {"question": "Workbook contains?",
             "options": ["Sheets", "Cells only", "Charts only", "Rows only"],
             "answer": 0},
        ]

        # -------------------------------------------------
        # Quiz State Variables
        # -------------------------------------------------
        self.total_questions = len(self.questions)
        self.current_question = 0
        self.score = 0
        self.answer_times = []
        self.question_start_time = None

        # -------------------------------------------------
        # UI Setup
        # -------------------------------------------------
        self.create_ui()
        self.show_question()

        # -------------------------------------------------
        # Proper Close Handling
        # -------------------------------------------------
        self.protocol("WM_DELETE_WINDOW", self.close_application)


    # =====================================================
    # WINDOW CENTERING (ONLY centers window — nothing else)
    # =====================================================

    def center_window(self, width, height):
        self.update_idletasks()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        self.geometry(f"{width}x{height}+{x}+{y}")

    # =====================================================
    # DATABASE + AUTO MIGRATION
    # =====================================================

    def init_database(self):
        self.conn = sqlite3.connect("quiz_results.db")
        self.cursor = self.conn.cursor()

        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT,
            score INTEGER,
            total INTEGER,
            percentage INTEGER,
            date TEXT
        )
        """)

        self.cursor.execute("PRAGMA table_info(results)")
        columns = [col[1] for col in self.cursor.fetchall()]

        if "avg_time" not in columns:
            self.cursor.execute("ALTER TABLE results ADD COLUMN avg_time REAL")

        if "total_time" not in columns:
            self.cursor.execute("ALTER TABLE results ADD COLUMN total_time REAL")

        self.conn.commit()

    # =====================================================
    # UI
    # =====================================================

    def create_ui(self):

        self.progress_label = ctk.CTkLabel(self, text="", font=("Arial", 16))
        self.progress_label.pack(pady=10)

        self.progress_bar = ctk.CTkProgressBar(self, width=800)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

        self.question_label = ctk.CTkLabel(
            self, text="", wraplength=850, font=("Arial", 26))
        self.question_label.pack(pady=30)

        self.buttons = []
        for i in range(4):
            btn = ctk.CTkButton(
                self,
                text="",
                width=800,
                height=55,
                font=("Arial", 16),
                command=lambda i=i: self.check_answer(i)
            )
            btn.pack(pady=10)
            self.buttons.append(btn)

        self.feedback_label = ctk.CTkLabel(self, text="", font=("Arial", 18))
        self.feedback_label.pack(pady=10)

        self.next_button = ctk.CTkButton(
            self, text="Next",
            command=self.next_question,
            state="disabled",
            width=200)
        self.next_button.pack(pady=10)

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(pady=20)

        ctk.CTkButton(bottom, text="Restart",
                      command=self.restart_quiz).grid(row=0, column=0, padx=10)

        ctk.CTkButton(bottom, text="Dashboard",
                      command=self.open_login_window).grid(row=0, column=1, padx=10)
        
        ctk.CTkButton(bottom,
              text="Export to Excel",
              command=self.export_to_excel).grid(row=0, column=2, padx=10)
         
        ctk.CTkButton(bottom,
            text="Export to PDF",
            command=self.export_to_pdf).grid(row=0, column=3, padx=10)
        
        ctk.CTkButton(bottom, text="Close",
                      fg_color="red",
                      command=self.close_application).grid(row=0, column=4, padx=10)
      
        
    # =====================================================
    # QUIZ LOGIC
    # =====================================================

    def show_question(self):
        if self.current_question < self.total_questions:

            q = self.questions[self.current_question]

            self.progress_label.configure(
                text=f"Question {self.current_question+1} of {self.total_questions}")

            self.progress_bar.set(
                self.current_question/self.total_questions)

            self.question_label.configure(text=q["question"])
            self.feedback_label.configure(text="")

            for i, option in enumerate(q["options"]):
                self.buttons[i].configure(
                    text=option,
                    state="normal"
                )

            self.next_button.configure(state="disabled")
            self.question_start_time = time.time()
        else:
            self.show_result()

    def check_answer(self, selected):
        time_taken = time.time() - self.question_start_time
        self.answer_times.append(round(time_taken, 2))

        correct = self.questions[self.current_question]["answer"]

        for btn in self.buttons:
            btn.configure(state="disabled")

        if selected == correct:
            self.score += 1
            self.feedback_label.configure(text="✅ Correct!", text_color="green")
        else:
            self.feedback_label.configure(text="❌ Wrong!", text_color="red")

        self.next_button.configure(state="normal")

    def next_question(self):
        self.current_question += 1
        self.show_question()

    def show_result(self):
        percentage = int((self.score/self.total_questions)*100)
        total_time = round(sum(self.answer_times), 2)
        
        avg_time = round(total_time/self.total_questions, 2) if self.total_questions else 0

        self.avg_time = avg_time
        self.total_time = total_time

        self.question_label.configure(
            text=f"""Quiz Completed!

Score: {self.score}/{self.total_questions}
Percentage: {percentage}%

Average Time per Question: {avg_time} sec
Total Time: {total_time} sec"""
        )

        for btn in self.buttons:
            btn.pack_forget()

        self.next_button.pack_forget()
        self.progress_bar.set(1)

        self.after(500, self.ask_save_to_database)

    # =====================================================
    # DUPLICATE EMAIL IMPROVED
    # =====================================================

    def ask_save_to_database(self):

        save = ctk.CTkToplevel(self)
        save.title("Save Result")
        save.geometry("500x420")
        save.transient(self)
        save.grab_set()

        ctk.CTkLabel(save, text="Save Results?",
                     font=("Arial", 22)).pack(pady=20)

        name_entry = ctk.CTkEntry(save,
                                  placeholder_text="Full Name",
                                  width=350,
                                  height=45)
        name_entry.pack(pady=10)

        email_entry = ctk.CTkEntry(save,
                                   placeholder_text="Email",
                                   width=350,
                                   height=45)
        email_entry.pack(pady=10)

        def save_result():

            name = name_entry.get()
            email = email_entry.get()

            if not name or not email:
                msg.showerror("Error", "Name and Email required")
                return

            percentage = int((self.score/self.total_questions)*100)
            today = datetime.now().strftime("%Y-%m-%d")

            self.cursor.execute("SELECT * FROM results WHERE email=?", (email,))
            existing = self.cursor.fetchall()

            if existing:
                choice = msg.askyesnocancel(
                    "Duplicate Email Found",
                    "This email already exists.\n\n"
                    "YES = Overwrite ALL previous attempts\n"
                    "NO = Add as new attempt\n"
                    "CANCEL = Do nothing"
                )

                if choice is None:
                    return
                elif choice:
                    self.cursor.execute("DELETE FROM results WHERE email=?", (email,))
                # else -> keep old + add new

            self.cursor.execute("""
            INSERT INTO results
            (name,email,score,total,percentage,avg_time,total_time,date)
            VALUES (?,?,?,?,?,?,?,?)
            """, (name, email,
                  self.score,
                  self.total_questions,
                  percentage,
                  self.avg_time,
                  self.total_time,
                  today))

            self.conn.commit()
            msg.showinfo("Saved", "Result saved successfully")
            save.destroy()

        ctk.CTkButton(save, text="Save",
                      width=200, height=45,
                      command=save_result).pack(pady=20)
        
        

    # =====================================================
    # PROFESSIONAL DASHBOARD
    # =====================================================

    def open_login_window(self):

        login = ctk.CTkToplevel(self)
        login.title("Admin Login")
        login.geometry("400x300")
        login.transient(self)
        login.grab_set()

        user = ctk.CTkEntry(login, placeholder_text="Username")
        user.pack(pady=20)

        pwd = ctk.CTkEntry(login, placeholder_text="Password", show="*")
        pwd.pack(pady=10)

        def check():
            if user.get() == ADMIN_USERNAME and pwd.get() == ADMIN_PASSWORD:
                login.destroy()
                self.open_dashboard()
            else:
                msg.showerror("Error", "Invalid credentials")

        ctk.CTkButton(login, text="Login",
                      command=check).pack(pady=20)

    def open_dashboard(self):

        dash = ctk.CTkToplevel(self)
        dash.title("Professional Results Dashboard")
        dash.geometry("1200x850")

        search_frame = ctk.CTkFrame(dash)
        search_frame.pack(pady=10)

        search_entry = ctk.CTkEntry(search_frame,
                                    placeholder_text="Search name or email",
                                    width=300)
        search_entry.grid(row=0, column=0, padx=10)

        table_frame = ctk.CTkScrollableFrame(dash, width=1150, height=350)
        table_frame.pack(pady=10)

        
        # --- BUTTON SECTION HERE ---
        bottom = ctk.CTkFrame(dash)
        bottom.pack(pady=15)

        ctk.CTkButton(bottom,
                    text="Export to Excel",
                    command=self.export_to_excel).grid(row=0, column=0, padx=10)

        ctk.CTkLabel(bottom,
                    text="Delete Record by ID:").grid(row=0, column=1, padx=10)

        delete_entry = ctk.CTkEntry(bottom, width=100)
        delete_entry.grid(row=0, column=2, padx=5)

        def delete_record():
            record_id = delete_entry.get().strip()

            if not record_id.isdigit():
                msg.showerror(
                    "Error",
                    "Please enter a valid numeric ID.",
                    parent=dash
                )
                delete_entry.focus()
                return

            confirm = msg.askyesno(
                "Confirm Delete",
                f"Delete record ID {record_id}?",
                parent=dash
            )

            if confirm:
                self.cursor.execute(
                    "DELETE FROM results WHERE id=?",
                    (record_id,)
                )
                self.conn.commit()

                msg.showinfo(
                    "Deleted",
                    "Record deleted successfully.",
                    parent=dash
                )

                delete_entry.delete(0, "end")
                delete_entry.focus()
                load_data()






        ctk.CTkButton(bottom,
                    text="Delete",
                    fg_color="red",
                    command=delete_record).grid(row=0, column=3, padx=10)
        #--

        # Scrollable analytics section
        analytics_frame = ctk.CTkScrollableFrame(dash, width=1150, height=350)
        analytics_frame.pack(pady=20, fill="both", expand=True)

        headers = ["ID", "Name", "Email", "Score", "Total",
                "Percentage", "Avg Time", "Total Time", "Date"]

        def load_data(filter_text=""):

            for widget in table_frame.winfo_children():
                widget.destroy()         

            for widget in analytics_frame.winfo_children():
                widget.destroy()    

            query = "SELECT * FROM results"
            params = ()

            if filter_text:
                query += " WHERE name LIKE ? OR email LIKE ?"
                params = (f"%{filter_text}%", f"%{filter_text}%")

            self.cursor.execute(query, params)
            records = self.cursor.fetchall()

            # Table Headers
            for col, h in enumerate(headers):
                ctk.CTkLabel(table_frame,
                            text=h,
                            font=("Arial", 14, "bold")
                            ).grid(row=0, column=col, padx=10)

            # Table Rows
            for r, row in enumerate(records, start=1):
                for c, val in enumerate(row):
                    ctk.CTkLabel(table_frame,
                                text=str(val)
                                ).grid(row=r, column=c, padx=10, pady=5)

            # =======================
            # EMBEDDED CHART SECTION
            # =======================

            if records:

                scores = [int(r[3]) for r in records]
                percentages = [int(r[5]) for r in records]
                avg_times = [float(r[6] or 0) for r in records]
                attempts = list(range(1, len(records)+1))

                plt.style.use("dark_background")
                fig, ax = plt.subplots(figsize=(10, 4))

                ax.plot(attempts, scores, marker='o', label="Score")
                ax.plot(attempts, percentages, marker='o', label="Percentage")
                ax.set_xticks(attempts)
                ax.set_yticks(range(0, 101, 10))
                ax.set_title("Performance Overview")
                ax.set_xlabel("Attempt")
                ax.legend()

                fig.tight_layout()

                canvas = FigureCanvasTkAgg(fig, master=analytics_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(fill="both", expand=True)

                plt.close(fig)   # <-- ADD THIS LINE HERE

            # =======================
            # GLOBAL STATS
            # =======================

            self.cursor.execute("""
            SELECT COUNT(*),
                AVG(score),
                AVG(percentage),
                AVG(avg_time),
                AVG(total_time),
                MAX(score),
                MIN(avg_time)
            FROM results
            """)
            stats = self.cursor.fetchone()

            stats_frame = ctk.CTkFrame(analytics_frame)
            stats_frame.pack(pady=25, fill="x")

            ctk.CTkLabel(stats_frame,
                        text="📊 Global Statistics",
                        font=("Arial", 18, "bold")
                        ).pack()

            ctk.CTkLabel(stats_frame,
                        text=f"""
    Total Attempts: {stats[0] or 0}

    Average Score: {round(stats[1] or 0,2)}
    Average Percentage: {round(stats[2] or 0,2)}%

    Average Time per Question: {round(stats[3] or 0,2)} sec
    Average Total Time: {round(stats[4] or 0,2)} sec

    Highest Score Achieved: {stats[5] or 0}
    Fastest Avg Time: {round(stats[6] or 0,2)} sec
    """,
                        justify="left").pack()

        load_data()

        search_entry.bind("<KeyRelease>",
                        lambda e: load_data(search_entry.get()))

             
              

    # =====================================================

    def export_to_excel(self):

        # Fetch all records
        self.cursor.execute("SELECT * FROM results")
        records = self.cursor.fetchall()

        if not records:
            msg.showinfo("No Data", "No records found to export.")
            return

        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Quiz Results"

        # Headers
        headers = ["ID", "Name", "Email", "Score", "Total",
                "Percentage", "Avg Time", "Total Time", "Date"]

        ws.append(headers)

        # Add data rows
        for row in records:
            ws.append(row)

        # Auto column width
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column].width = adjusted_width

        # Save file
        filename = f"quiz_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(filename)

        msg.showinfo("Success", f"Results exported successfully!\n\nSaved as:\n{filename}")

    # =====================================================

    def export_to_pdf(self):

        self.cursor.execute("SELECT * FROM results")
        records = self.cursor.fetchall()

        if not records:
            msg.showinfo("No Data", "No records found to export.")
            return

        filename = f"quiz_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        doc = SimpleDocTemplate(
            filename,
            pagesize=A4
        )

        elements = []
        style_sheet = styles.getSampleStyleSheet()

        # =====================================================
        # TITLE SECTION (kept together)
        # =====================================================

        title_section = []
        title_section.append(Paragraph("Excel Quiz Results Report - By Tryfon Papadopoulos", style_sheet["Heading1"]))
        title_section.append(Spacer(1, 0.3 * inch))

        elements.append(KeepTogether(title_section))

        # =====================================================
        # TABLE SECTION (kept together with title)
        # =====================================================

        headers = ["ID", "Name", "Email", "Score", "Total",
                "Percentage", "Avg Time", "Total Time", "Date"]

        data = [headers]
        for row in records:
            data.append([str(cell) for cell in row])

        table = Table(data, repeatRows=1)

        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (3, 1), (6, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ]))

        table_section = [
            Paragraph("Results Table", style_sheet["Heading2"]),
            Spacer(1, 0.2 * inch),
            table,
            Spacer(1, 0.4 * inch)
        ]

        elements.append(KeepTogether(table_section))

        # =====================================================
        # CHART SECTION (kept together)
        # =====================================================

        scores = [int(r[3]) for r in records]
        percentages = [int(r[5]) for r in records]
        attempts = list(range(1, len(records) + 1))

        fig, ax = plt.subplots(figsize=(6, 3))
        ax.plot(attempts, scores, marker='o', label="Score")
        ax.plot(attempts, percentages, marker='o', label="Percentage")
        ax.set_xticks(attempts)
        ax.set_yticks(range(0, 101, 10))
        ax.set_title("Performance Overview")
        ax.set_xlabel("Attempt")
        ax.legend()

        fig.tight_layout()

        chart_filename = "temp_chart.png"
        fig.savefig(chart_filename, dpi=200)
        plt.close(fig)

        chart_section = [
            Paragraph("Performance Chart", style_sheet["Heading2"]),
            Spacer(1, 0.2 * inch),
            Image(chart_filename, width=6 * inch, height=3 * inch),
            Spacer(1, 0.4 * inch)
        ]

        elements.append(KeepTogether(chart_section))

        # =====================================================
        # SUMMARY SECTION (kept together)
        # =====================================================

        self.cursor.execute("""
            SELECT COUNT(*),
                AVG(score),
                AVG(percentage),
                AVG(avg_time),
                AVG(total_time),
                MAX(score),
                MIN(avg_time)
            FROM results
        """)
        stats = self.cursor.fetchone()

        summary_text = f"""
    Total Attempts: {stats[0] or 0}<br/>
    Average Score: {round(stats[1] or 0,2)}<br/>
    Average Percentage: {round(stats[2] or 0,2)}%<br/>
    Average Time per Question: {round(stats[3] or 0,2)} sec<br/>
    Average Total Time: {round(stats[4] or 0,2)} sec<br/>
    Highest Score: {stats[5] or 0}<br/>
    Fastest Avg Time: {round(stats[6] or 0,2)} sec
    """

        summary_section = [
            Paragraph("Summary Statistics", style_sheet["Heading2"]),
            Spacer(1, 0.2 * inch),
            Paragraph(summary_text, style_sheet["Normal"])
        ]

        elements.append(KeepTogether(summary_section))

    
        # =====================================================
        # PAGE NUMBER FUNCTION
        # =====================================================

        def add_page_number(canvas, doc):
            page_num_text = f"Page {canvas.getPageNumber()}"
            canvas.setFont("Helvetica", 9)
            canvas.drawRightString(doc.pagesize[0] - 40, 20, page_num_text)

        # Build PDF with page numbers
        doc.build(
            elements,
            onFirstPage=add_page_number,
            onLaterPages=add_page_number
        )
     
        # Remove temporary chart
        if os.path.exists(chart_filename):
            os.remove(chart_filename)

        msg.showinfo("Success", f"PDF report exported successfully!\n\nSaved as:\n{filename}")

    def close_application(self):
        try:
            plt.close('all')   # <-- VERY IMPORTANT (closes matplotlib)
        except:
            pass

        try:
            self.conn.close()
        except:
            pass

        self.after(100, self.destroy)
    
    
    def restart_quiz(self):
        self.current_question = 0
        self.score = 0
        self.answer_times = []
        for btn in self.buttons:
            btn.pack(pady=10)
        self.next_button.pack(pady=10)
        self.show_question()


if __name__ == "__main__":
    app = ExcelQuizApp()
    app.mainloop()