import tkinter as tk
from tkinter import messagebox
import pandas as pd

class StudentInfoSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Öğrenci Not Bilgi Sistemi")

        self.entries = []
        self.student_records = self.load_student_records()

        self.name_label = tk.Label(root, text="İsim:")
        self.name_label.pack()
        self.name_entry = tk.Entry(root)
        self.name_entry.pack()

        self.surname_label = tk.Label(root, text="Soyisim:")
        self.surname_label.pack()
        self.surname_entry = tk.Entry(root)
        self.surname_entry.pack()

        self.exam_number_label = tk.Label(root, text="Sınav Sayısı:")
        self.exam_number_label.pack()
        self.exam_number_entry = tk.Entry(root)
        self.exam_number_entry.pack()

        self.create_exams_button = tk.Button(root, text="Sınav Kutucuklarını Oluştur", command=self.create_exam_entries)
        self.create_exams_button.pack()

        self.save_button = tk.Button(root, text="Kaydet", command=self.save_student_record)
        self.save_button.pack()

    def create_exam_entries(self):
        try:
            exam_number = int(self.exam_number_entry.get())
        except ValueError:
            messagebox.showerror("Hata", "Geçerli bir sınav sayısı girin.")
            return

        for entry in self.entries:
            entry.destroy()

        self.entries = []

        for i in range(exam_number):
            label = tk.Label(self.root, text=f"{i+1}. Sınav:")
            label.pack()
            entry = tk.Entry(self.root)
            entry.pack()
            self.entries.append(entry)

    def save_student_record(self):
        student = {
            'name': self.name_entry.get(),
            'surname': self.surname_entry.get(),
            'grades': [],
            'average': 0
        }

        for entry in self.entries:
            try:
                grade = float(entry.get())
                student['grades'].append(grade)
            except ValueError:
                messagebox.showerror("Hata", "Geçerli bir not girin.")
                return

        if student['grades']:
            student['average'] = sum(student['grades']) / len(student['grades'])

        self.student_records.append(student)
        self.save_student_records()

    def save_student_records(self):
        df = pd.DataFrame(self.student_records)
        df.to_excel('D:\\Öğrencibilgisistemi.xlsx', index=False)

    def load_student_records(self):
        try:
            df = pd.read_excel('D:\\Öğrencibilgisistemi.xlsx')
            return df.to_dict('records')
        except (FileNotFoundError, ValueError):
            return []

if __name__ == '__main__':
    root = tk.Tk()
    app = StudentInfoSystem(root)
    root.mainloop()
