import pandas as pd
import tkinter as tk
from datetime import date, datetime
import os

class GymApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("King Abdulaziz University Gym")
        self.geometry("960x740")

        # Configure the theme
        self.configure(bg="#f0fff0")  # Set background color to green

        self.create_widgets()

    def create_widgets(self):
        self.logo_image = tk.PhotoImage(file="C:\\Users\\PCD GAMING\\Desktop\\Main Apps\\Programming Apps\\Projects\\Gym Attendance System\\KAU_logo.png")
        self.logo_label = tk.Label(self, image=self.logo_image, bg="#f0fff0")  # Set background color to green
        self.logo_label.pack()

        self.title_label = tk.Label(self, text="النادي الرياضي بجامعة الملك عبدالعزيز",
                                    font=("Arial", 16, "bold"), bg="#f0fff0", fg="green")  # Set text color to green
        self.title_label.pack(pady=10)

        self.prompt_label = tk.Label(self, text="الرجاء ادخال رقم الجوال", font=("Arial", 12), bg="#f0fff0")
        self.prompt_label.pack()

        self.response = tk.StringVar()

        self.response_entry = tk.Entry(self, textvariable=self.response, font=("Arial", 12))
        self.response_entry.pack()

        self.submit_button = tk.Button(self, text="تسجيل", font=("Arial", 12), bg="green", fg="#f0fff0")  # Set button color to green
        self.submit_button.pack(pady=10)

        self.response_label = tk.Label(self, text="", font=("Arial", 12), bg="#f0fff0")
        self.response_label.pack()

        # Bind the Enter key event to the submit_primary_key method
        self.response_entry.bind("<Return>", self.submit_primary_key)

    def submit_primary_key(self, event):
        primary_key = self.response.get()

        # Read the Excel file
        df = pd.read_excel("C:\\Users\\PCD GAMING\\Desktop\\Main Apps\\Programming Apps\\Projects\\Gym Attendance System\\GYM_DATA.xlsx", sheet_name="DATA")

        # Check if the primary key exists in the dataframe
        if int(primary_key) in df["Primary Key"].values:

            # Get the user name
            user_name = df.loc[df["Primary Key"] == int(primary_key), "User Name"].values[0]

            self.response_label.config(text=f"تم تحضيرك بنجاح. اسم المستخدم: {user_name}", fg="green")

            # Save the attendance with the date and time
            now = datetime.now().strftime("%H:%M:%S")
            attendance_data = pd.DataFrame({"User Name": [user_name], "Date": [date.today()], "Time": [now]})

            # Write the updated data to the file
            output_file = "C:\\Users\\PCD GAMING\\Desktop\\Main Apps\\Programming Apps\\Projects\\Gym Attendance System\\new_attendees.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                start_row = writer.sheets["Attendance"].max_row + 0
                attendance_data.to_excel(writer, sheet_name="Attendance", startrow=start_row, index=False)




        else:
            self.response_label.config(text="خطا, حاول التاكد من انك قمت بتسجيل رقم الجوال الصحيح.", fg="red")


if __name__ == "__main__":
    app = GymApp()
    app.mainloop()
