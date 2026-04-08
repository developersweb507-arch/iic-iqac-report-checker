import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import os
import re
from datetime import datetime
import tempfile

try:
    import PyPDF2
    pdf_support = True
except:
    pdf_support = False

try:
    import docx
    docx_support = True
except:
    docx_support = False

try:
    from pdf2image import convert_from_path
    import pytesseract
    ocr_support = True
    
    # PATH SET KAR DIYA HAI - TUMHARE FOLDER KE ACCORDING
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    # POPPLER PATH - TUMHARE DOWNLOADS FOLDER MEIN
    POPPLER_PATH = r"C:\Users\pc\Downloads\report checker\poppler-25.12.0\Library\bin"
    
except:
    ocr_support = False
    POPPLER_PATH = None

class ReportCheckerFinal:
    def __init__(self, root):
        self.root = root
        self.root.title("IIC/IQAC Report Checker - Complete")
        self.root.geometry("950x750")
        self.root.configure(bg='#f0f0f0')
        
        self.report_text = ""
        self.file_path = ""
        self.file_loaded = False
        
        # Mandatory checks
        self.mandatory_checks = {
            "1": "Notice (Program name, Date, Venue, Coordinator/HOD)",
            "2": "Poster/Banner of the activity/event",
            "3": "Schedule (Day-wise/Hour-wise)",
            "4": "Write-up (Theme, Objective, Summary, Participants)",
            "5": "Expert/Speaker details (Name, Organisation, Designation)",
            "6": "Social media links (Insta, FB, LinkedIn, Twitter)",
            "7": "Video URL (YouTube link)",
            "8": "Sample certificate",
            "9": "Attendance record",
            "10": "Feedback form and analysis",
            "11": "Geotagged photographs (Minimum 4 photos)",
            "12": "Email to iic@ssipmt.com",
            "13": "Hard copy submitted to IQAC (within 2 days)",
            "14": "Remuneration proof (Cheque/NEFT copy)"
        }
        
        self.keywords = {
            "1": ["notice", "program", "date", "venue", "coordinator", "convener", "hod"],
            "2": ["poster", "banner", "flyer"],
            "3": ["schedule", "day-wise", "hour-wise", "timeline", "itinerary"],
            "4": ["write-up", "theme", "objective", "summary", "participants", "duration", "mode"],
            "5": ["expert", "speaker", "resource", "organisation", "designation"],
            "6": ["instagram", "facebook", "linkedin", "twitter", "social media"],
            "7": ["video", "youtube", "youtu.be"],
            "8": ["certificate", "sample certificate"],
            "9": ["attendance", "attendee", "participant list"],
            "10": ["feedback", "evaluation", "feedback form"],
            "11": ["geotag", "geotagged", "photo", "photograph", "caption", "4 photo"],
            "12": ["iic@ssipmt.com", "email", "mailed"],
            "13": ["hard copy", "iqac", "submitted", "2 days", "two days"],
            "14": ["cheque", "neft", "remuneration", "payment", "honorarium"]
        }
        
        self.setup_ui()
    
    def setup_ui(self):
        header = tk.Label(self.root, text="IIC & IQAC Report Compliance Checker", 
                         font=("Arial", 18, "bold"), bg='#2c3e50', fg='white', pady=15)
        header.pack(fill='x')
        
        subheader = tk.Label(self.root, text="Supports: Text PDF | Scanned PDF (OCR) | Word | Text Files", 
                            font=("Arial", 10), bg='#34495e', fg='white', pady=5)
        subheader.pack(fill='x')
        
        main_frame = tk.Frame(self.root, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)
        
        upload_frame = tk.Frame(main_frame, bg='#ecf0f1', relief=tk.RAISED, bd=2)
        upload_frame.pack(fill='x', pady=10)
        
        tk.Label(upload_frame, text="📁 Step 1: Upload Report File", 
                font=("Arial", 14, "bold"), bg='#ecf0f1', fg='#2c3e50').pack(pady=10)
        
        btn_frame = tk.Frame(upload_frame, bg='#ecf0f1')
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="📄 PDF (Normal)", command=self.load_pdf,
                 bg='#e74c3c', fg='white', font=("Arial", 10), padx=15, pady=8).pack(side='left', padx=5)
        
        tk.Button(btn_frame, text="🔍 PDF (Scanned/OCR)", command=self.load_pdf_ocr,
                 bg='#e67e22', fg='white', font=("Arial", 10), padx=15, pady=8).pack(side='left', padx=5)
        
        tk.Button(btn_frame, text="📝 Word (DOCX)", command=self.load_docx,
                 bg='#3498db', fg='white', font=("Arial", 10), padx=15, pady=8).pack(side='left', padx=5)
        
        tk.Button(btn_frame, text="📃 Text (TXT)", command=self.load_txt,
                 bg='#2ecc71', fg='white', font=("Arial", 10), padx=15, pady=8).pack(side='left', padx=5)
        
        self.file_label = tk.Label(upload_frame, text="No file selected", 
                                   font=("Arial", 10), bg='#ecf0f1', fg='#7f8c8d')
        self.file_label.pack(pady=5)
        
        self.status_label = tk.Label(upload_frame, text="⚪ Ready - Select a file to begin", 
                                     font=("Arial", 10), bg='#ecf0f1', fg='#e67e22')
        self.status_label.pack(pady=5)
        
        self.check_btn = tk.Button(main_frame, text="🔍 Step 2: Check Report", 
                                   command=self.check_report,
                                   bg='#2c3e50', fg='white', font=("Arial", 14, "bold"),
                                   padx=30, pady=10, state='disabled')
        self.check_btn.pack(pady=15)
        
        results_frame = tk.Frame(main_frame, bg='white', relief=tk.RAISED, bd=2)
        results_frame.pack(fill='both', expand=True, pady=10)
        
        tk.Label(results_frame, text="📊 Step 3: Compliance Report", 
                font=("Arial", 14, "bold"), bg='white', fg='#2c3e50').pack(pady=10)
        
        self.results_text = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, 
                                                       width=85, height=18,
                                                       font=("Courier", 10))
        self.results_text.pack(fill='both', expand=True, padx=10, pady=10)
        
        bottom_frame = tk.Frame(main_frame, bg='#f0f0f0')
        bottom_frame.pack(fill='x', pady=10)
        
        self.save_btn = tk.Button(bottom_frame, text="💾 Save as Text", 
                                  command=self.save_report,
                                  bg='#27ae60', fg='white', font=("Arial", 11),
                                  padx=15, pady=5, state='disabled')
        self.save_btn.pack(side='left', padx=5)
        
        tk.Button(bottom_frame, text="🖨️ Print", 
                 command=self.print_report,
                 bg='#9b59b6', fg='white', font=("Arial", 11),
                 padx=15, pady=5).pack(side='left', padx=5)
    
    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[["PDF files", "*.pdf"]])
        if file_path and pdf_support:
            self.file_path = file_path
            self.file_label.config(text=f"Loaded: {os.path.basename(file_path)}")
            self.status_label.config(text="⏳ Reading PDF...", fg='#3498db')
            self.root.update()
            
            try:
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    text = ""
                    for page in pdf_reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + " "
                    
                    if text.strip():
                        self.report_text = text.lower()
                        self.file_loaded = True
                        self.check_btn.config(state='normal')
                        self.status_label.config(text=f"✅ PDF loaded! {len(self.report_text)} chars", fg='#27ae60')
                        messagebox.showinfo("Success", f"PDF loaded!\nPages: {len(pdf_reader.pages)}\nText length: {len(self.report_text)} chars")
                    else:
                        self.status_label.config(text="❌ No text found! Use 'PDF (Scanned/OCR)' button", fg='#e74c3c')
            except Exception as e:
                self.status_label.config(text=f"❌ Error: {str(e)[:50]}", fg='#e74c3c')
    
    def load_pdf_ocr(self):
        if not ocr_support:
            messagebox.showerror("Error", "OCR packages not installed.\nRun: pip install pytesseract pdf2image")
            return
        
        file_path = filedialog.askopenfilename(filetypes=[["PDF files", "*.pdf"]])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"Loaded: {os.path.basename(file_path)} (OCR Mode)")
            self.status_label.config(text="⏳ OCR processing... This may take 2-3 minutes", fg='#e67e22')
            self.root.update()
            
            try:
                # Check if poppler exists
                if not os.path.exists(POPPLER_PATH):
                    self.status_label.config(text="❌ Poppler not found! Use Text File instead", fg='#e74c3c')
                    messagebox.showerror("Error", f"Poppler not found at:\n{POPPLER_PATH}\n\nPlease use 'Text File' option instead.")
                    return
                
                # Convert PDF to images
                self.status_label.config(text="⏳ Converting PDF to images...", fg='#e67e22')
                self.root.update()
                images = convert_from_path(file_path, dpi=200, poppler_path=POPPLER_PATH)
                
                text = ""
                total_pages = len(images)
                
                for i, image in enumerate(images, 1):
                    self.status_label.config(text=f"⏳ OCR page {i}/{total_pages}...", fg='#e67e22')
                    self.root.update()
                    page_text = pytesseract.image_to_string(image, lang='eng')
                    text += page_text + " "
                
                if text.strip():
                    self.report_text = text.lower()
                    self.file_loaded = True
                    self.check_btn.config(state='normal')
                    self.status_label.config(text=f"✅ OCR complete! {len(self.report_text)} chars", fg='#27ae60')
                    messagebox.showinfo("Success", f"OCR Completed!\nPages: {total_pages}\nText length: {len(self.report_text)} chars")
                else:
                    self.status_label.config(text="❌ No text recognized", fg='#e74c3c')
                    
            except Exception as e:
                self.status_label.config(text=f"❌ Error: {str(e)[:50]}", fg='#e74c3c')
                messagebox.showerror("Error", f"OCR failed: {str(e)}")
    
    def load_docx(self):
        file_path = filedialog.askopenfilename(filetypes=[["Word files", "*.docx"]])
        if file_path and docx_support:
            self.file_path = file_path
            self.file_label.config(text=f"Loaded: {os.path.basename(file_path)}")
            self.status_label.config(text="⏳ Reading Word file...", fg='#3498db')
            self.root.update()
            
            try:
                doc = docx.Document(file_path)
                text = ""
                for para in doc.paragraphs:
                    if para.text:
                        text += para.text + "\n"
                
                if text.strip():
                    self.report_text = text.lower()
                    self.file_loaded = True
                    self.check_btn.config(state='normal')
                    self.status_label.config(text=f"✅ Word loaded! {len(self.report_text)} chars", fg='#27ae60')
                    messagebox.showinfo("Success", f"Word file loaded!\nText length: {len(self.report_text)} chars")
                else:
                    self.status_label.config(text="❌ No text found!", fg='#e74c3c')
            except Exception as e:
                self.status_label.config(text=f"❌ Error: {str(e)[:50]}", fg='#e74c3c')
    
    def load_txt(self):
        file_path = filedialog.askopenfilename(filetypes=[["Text files", "*.txt"]])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"Loaded: {os.path.basename(file_path)}")
            self.status_label.config(text="⏳ Reading text file...", fg='#3498db')
            self.root.update()
            
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    text = file.read()
                
                if text.strip():
                    self.report_text = text.lower()
                    self.file_loaded = True
                    self.check_btn.config(state='normal')
                    self.status_label.config(text=f"✅ Text loaded! {len(self.report_text)} chars", fg='#27ae60')
                    messagebox.showinfo("Success", f"Text file loaded!\nText length: {len(self.report_text)} chars")
                else:
                    self.status_label.config(text="❌ File is empty!", fg='#e74c3c')
            except Exception as e:
                self.status_label.config(text=f"❌ Error: {str(e)[:50]}", fg='#e74c3c')
    
    def check_keywords(self, keywords):
        found = sum(1 for kw in keywords if kw in self.report_text)
        return found >= len(keywords) * 0.5
    
    def check_geotag_photos(self):
        if 'geotag' in self.report_text:
            return True
        numbers = re.findall(r'(\d+)\s*photos?', self.report_text)
        for num in numbers:
            if int(num) >= 4:
                return True
        return False
    
    def check_report(self):
        if not self.file_loaded or not self.report_text:
            messagebox.showwarning("Warning", "Please load a report file first!")
            return
        
        self.results_text.delete(1.0, tk.END)
        
        present_items = []
        missing_items = []
        
        self.results_text.insert(tk.END, "="*80 + "\n")
        self.results_text.insert(tk.END, "📊 IIC/IQAC REPORT COMPLIANCE VERIFICATION\n")
        self.results_text.insert(tk.END, "="*80 + "\n")
        self.results_text.insert(tk.END, f"📅 Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.results_text.insert(tk.END, f"📄 File: {os.path.basename(self.file_path)}\n")
        self.results_text.insert(tk.END, f"📏 Text Length: {len(self.report_text)} characters\n")
        self.results_text.insert(tk.END, "="*80 + "\n\n")
        
        for i in range(1, 15):
            key = str(i)
            if key == "11":
                is_present = self.check_geotag_photos()
            else:
                is_present = self.check_keywords(self.keywords[key])
            
            if is_present:
                present_items.append(key)
                status = "✅ PASS"
            else:
                missing_items.append(key)
                status = "❌ FAIL"
            
            self.results_text.insert(tk.END, f"{status}  {i}. {self.mandatory_checks[key]}\n")
        
        total = len(self.mandatory_checks)
        passed = len(present_items)
        score = (passed / total) * 100
        
        self.results_text.insert(tk.END, "\n" + "="*80 + "\n")
        self.results_text.insert(tk.END, "📈 SUMMARY\n")
        self.results_text.insert(tk.END, "="*80 + "\n")
        self.results_text.insert(tk.END, f"✅ Items Present: {passed}/{total} ({score:.1f}%)\n")
        self.results_text.insert(tk.END, f"❌ Items Missing: {len(missing_items)}/{total}\n")
        
        if missing_items:
            self.results_text.insert(tk.END, "\n" + "="*80 + "\n")
            self.results_text.insert(tk.END, "⚠️ MISSING ITEMS (INHE JODNA HAI):\n")
            self.results_text.insert(tk.END, "="*80 + "\n")
            for item in missing_items:
                self.results_text.insert(tk.END, f"  • {self.mandatory_checks[item]}\n")
        
        self.results_text.insert(tk.END, "\n" + "="*80 + "\n")
        if score == 100:
            self.results_text.insert(tk.END, "🎉 EXCELLENT! Report is 100% compliant!\n")
        elif score >= 80:
            self.results_text.insert(tk.END, f"⚠️ GOOD! Minor issues to fix. Score: {score:.1f}%\n")
        elif score >= 60:
            self.results_text.insert(tk.END, f"⚠️ SATISFACTORY - Needs improvement. Score: {score:.1f}%\n")
        else:
            self.results_text.insert(tk.END, f"❌ POOR - Major revisions required. Score: {score:.1f}%\n")
        self.results_text.insert(tk.END, "="*80 + "\n")
        
        self.save_btn.config(state='normal')
    
    def save_report(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", 
                                                  filetypes=[["Text files", "*.txt"]])
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.results_text.get(1.0, tk.END))
            messagebox.showinfo("Success", f"Report saved to:\n{file_path}")
    
    def print_report(self):
        try:
            fd, path = tempfile.mkstemp(suffix='.txt')
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write(self.results_text.get(1.0, tk.END))
            os.startfile(path, 'print')
        except Exception as e:
            messagebox.showerror("Error", f"Cannot print: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportCheckerFinal(root)
    root.mainloop()