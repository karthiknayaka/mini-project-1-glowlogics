from tkinter import *
from tkinter import messagebox
import xlsxwriter
import os
import mysql.connector

twindow = Tk()
twindow.geometry("560x450+300+100")

os.makedirs("C:\\Training", exist_ok=True)
file_path = "C:\\Training\\cambridge.xlsx"

# Labels
Label(twindow, text="STUDENT DETAILS", fg="purple", font=('broadway', 16)).place(x=145, y=20)

# Entry fields
Label(twindow, text="ROLLNO").place(x=20, y=80)
stf1 = Entry(twindow); stf1.place(x=200, y=80)

Label(twindow, text="NAME").place(x=20, y=110)
stf2 = Entry(twindow); stf2.place(x=200, y=110)

Label(twindow, text="SUB1").place(x=20, y=140)
stf3 = Entry(twindow); stf3.place(x=200, y=140)

Label(twindow, text="SUB2").place(x=20, y=170)
stf4 = Entry(twindow); stf4.place(x=200, y=170)

Label(twindow, text="SUB3").place(x=20, y=200)
stf5 = Entry(twindow); stf5.place(x=200, y=200)

Label(twindow, text="TOTAL").place(x=20, y=230)
stf6 = Entry(twindow); stf6.place(x=200, y=230)

Label(twindow, text="AVERAGE").place(x=20, y=260)
stf7 = Entry(twindow); stf7.place(x=200, y=260)


# Common calculation
def get_data():
    rollno = stf1.get()
    name = stf2.get()
    s1 = int(stf3.get())
    s2 = int(stf4.get())
    s3 = int(stf5.get())

    total = s1 + s2 + s3
    avg = total / 3

    stf6.delete(0, END)
    stf7.delete(0, END)
    stf6.insert(0, total)
    stf7.insert(0, round(avg, 2))

    return rollno, name, s1, s2, s3, total, avg


#  EXCEL BUTTON FUNCTION
def save_excel():
    try:
        data = get_data()

        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet()

        headers = ["RollNo", "Name", "Sub1", "Sub2", "Sub3", "Total", "Average"]
        worksheet.write_row(0, 0, headers)
        worksheet.write_row(1, 0, data)

        workbook.close()

        # PRINT IN TERMINAL
        print("EXCEL DATA SAVED:")
        print(data)

        messagebox.showinfo("SUCCESS", "Saved to Excel")

    except Exception as e:
        messagebox.showerror("ERROR", str(e))


#  MYSQL BUTTON FUNCTION
def save_mysql():
    try:
        data = get_data()

        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="1234"   # change your password
        )

        cursor = conn.cursor()

        # Create DB
        cursor.execute("CREATE DATABASE IF NOT EXISTS student_db")
        cursor.execute("USE student_db")

        # Create Table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS students(
            rollno VARCHAR(10),
            name VARCHAR(50),
            sub1 INT,
            sub2 INT,
            sub3 INT,
            total INT,
            average FLOAT
        )
        """)

        # Insert Data
        cursor.execute(
            "INSERT INTO students VALUES (%s,%s,%s,%s,%s,%s,%s)", data
        )

        conn.commit()

        # PRINT IN TERMINAL
        print("MYSQL DATA SAVED:")
        print(data)

        # Display all data
        cursor.execute("SELECT * FROM students")
        rows = cursor.fetchall()
        print("DATABASE CONTENT:")
        for r in rows:
            print(r)

        messagebox.showinfo("SUCCESS", "Saved to MySQL")

    except Exception as e:
        messagebox.showerror("ERROR", str(e))


# Buttons
Button(twindow, text="EXCEL", bg="green", fg="white",
       command=save_excel).place(x=120, y=320)

Button(twindow, text="MYSQL", bg="blue", fg="white",
       command=save_mysql).place(x=300, y=320)


twindow.mainloop()
