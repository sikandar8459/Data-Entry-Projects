from tkinter import *
from tkinter import ttk
import tkinter.messagebox as msg
import os
import openpyxl
import sqlite3

conn = sqlite3.connect("DataEntryDatabase.db")
cur = conn.cursor()

class DataEntryClass:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Entry Form | Developed by Sikandar Singh")
        self.root.focus_force()
        self.root.config(bg = "white")
        
        self.first_name = StringVar()
        self.last_name = StringVar()
        self.title = StringVar()

        self.gender = StringVar()
        self.age = StringVar()
        self.address = StringVar()
        
        self.information_frame = LabelFrame(self.root, text = "Information Details", bd = 2, relief = GROOVE, bg = "white")
        self.information_frame.pack(padx = 10, pady = 10)
        
        Label(self.information_frame, bg = "white", text = "First Name").grid(row = 0, column = 0, sticky = "w")
        Entry(self.information_frame, bd = 1, relief = GROOVE, textvariable = self.first_name).grid(row = 1, column = 0, sticky = "w")
        
        Label(self.information_frame, bg = "white", text = "Last Name").grid(row = 0, column = 1, sticky = "w")
        Entry(self.information_frame, bd = 1, relief = GROOVE, textvariable = self.last_name).grid(row = 1, column = 1, sticky = "w")
        
        Label(self.information_frame, bg = "white", text = "Title").grid(row = 0, column = 2, sticky = "w")
        self.title_combo = ttk.Combobox(self.information_frame, state = "readonly", textvariable = self.title, values = ["Select","Mr.","Ms.","Dr.","Eng."])
        self.title_combo.grid(row = 1, column = 2, sticky = "w")
        self.title.set("Select")

        Label(self.information_frame, bg = "white", text = "Gender").grid(row = 2, column = 0, sticky = "w")
        Entry(self.information_frame, bd = 1, relief = GROOVE, textvariable = self.gender).grid(row = 3, column = 0, sticky = "w")
        
        Label(self.information_frame, bg = "white", text = "Age").grid(row = 2, column = 1, sticky = "w")
        self.spin_age = Spinbox(self.information_frame, state = "readonly", textvariable = self.age, from_ = 18, to = 110)
        self.spin_age.grid(row = 3, column = 1)
        
        Label(self.information_frame, bg = "white", text = "Address").grid(row = 2, column = 2, sticky = "w")
        Entry(self.information_frame, bd = 1, relief = GROOVE, textvariable = self.address).grid(row = 3, column = 2, sticky = "w")

        for widget in self.information_frame.winfo_children():
            widget.grid_configure(padx = 10, pady = 10)
    
        self.courses_frame = LabelFrame(self.root, text = "Courses Details", bd = 2, relief = GROOVE, bg = "white")
        self.courses_frame.pack(padx = 10, pady = 10, fill = X)

        Label(self.courses_frame, bg = "white", text = "Course Name").grid(row = 0, column = 0, sticky = "w")
        Label(self.courses_frame, bg = "white", text = "Course Duration").grid(row = 0, column = 1, sticky = "w")
        Label(self.courses_frame, bg = "white", text = "Course Fees").grid(row = 0, column = 2, sticky = "w")

        self.course_name = StringVar()
        self.course_duration = StringVar()
        self.course_fees = StringVar()
        
        Entry(self.courses_frame, bd = 1, relief = GROOVE, textvariable = self.course_name).grid(row = 1, column = 0, sticky = "w")
        Entry(self.courses_frame, bd = 1, relief = GROOVE, textvariable = self.course_duration).grid(row = 1, column = 1, sticky = "w")
        Entry(self.courses_frame, bd = 1, relief = GROOVE, textvariable = self.course_fees).grid(row = 1, column = 2, sticky = "w")

        for widget in self.courses_frame.winfo_children():
            widget.grid_configure(padx = 10, pady = 10)
    
        self.terms_frame = LabelFrame(self.root, text = "Courses Details", bd = 2, relief = GROOVE, bg = "white")
        self.terms_frame.pack(padx = 10, pady = 10, expand = 1, fill = X)

        self.termsValue = StringVar()
        self.termsCheck = Checkbutton(self.terms_frame, variable = self.termsValue, text = "I accept the terms & conditions", activebackground = "white", cursor = "hand2", bg = "white", onvalue = "Checked", offvalue = 'Not Checked')
        self.termsCheck.grid(row = 0, column = 0, sticky = "w")
        self.termsValue.set("Not Checked")
    
        for widget in self.terms_frame.winfo_children():
            widget.grid_configure(padx = 10, pady = 10)
            
        Button(self.root, text = "Data Entry", cursor = "hand2", command = self.SubmitData).pack(padx = 10, pady = 10, fill = X)
        self.save_excel()
    
    def SubmitData(self):
        if self.termsValue.get() == "Not Checked":
            msg.showerror("Error","Please accept the terms & conditions...", parent = self.root)
        else:
            if self.first_name.get() == "" or self.last_name.get() == "" or self.title.get() == "" or self.gender.get() == "" or self.age.get() == "17" or self.address.get() == "" or self.course_name.get() == "" or self.course_duration.get() == "" or self.course_fees.get() == "":
                msg.showerror("Error","All fields are required..", parent = self.root)
            else:
                cur.execute("insert into DataEntryTable values (?,?,?,?,?,?,?,?,?,?)",(
                        self.first_name.get(),
                        self.last_name.get(),
                        self.title.get(),
                        self.gender.get(),
                        self.age.get(),
                        self.address.get(),
                        self.course_name.get(),
                        self.course_duration.get(),
                        self.course_fees.get(),
                        self.termsValue.get(),
                    ))
                conn.commit()
                msg.showinfo("Success","Data has been Added..", parent = self.root)
                self.save_excel()
                self.clear()
                
    def clear(self):
        self.first_name.set("")
        self.last_name.set("")
        self.title.set("Select")
        self.gender.set("")
        self.age.set("18")
        self.address.set("")
        self.course_name.set("")
        self.course_duration.set("")
        self.course_fees.set("")
        self.termsValue.set("Not Checked")
        
    def save_excel(self):
        self.filepath = "D:\Sikandar Singh\Python Programming\Data Entry Project\Information_Details_Sheet.xlsx"
        if os.path.exists(self.filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            heading = ["First Name", "Last Name", "Title", "Gender", "Age", "Address", "Course Name", "Course Duration", "Course Fees", "Terms & Conditions"]
            sheet.append(heading)
            workbook.save(self.filepath)
        workbook = openpyxl.load_workbook(self.filepath)
        sheet = workbook.active
        cur.execute("select * from DataEntryTable")
        result = cur.fetchall()
        for row in result:
            sheet.append(row)
            workbook.save(self.filepath)
                    
if __name__ == "__main__":
    root = Tk()
    obj = DataEntryClass(root)
    root.mainloop()