import customtkinter
from createreport import create_report, load_IDQ, load_VOL

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.title("WBR Report Creation Tool")
root.geometry("300x696")

# Frame for loading buttons
frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=10, padx=10, fill="both", expand=False)

# Frame for create button
button_frame = customtkinter.CTkFrame(master=root)
button_frame.pack(pady=0, padx=10, fill="both", expand=False)

# Frame for feedback info
info_frame = customtkinter.CTkFrame(master=root)
info_frame.pack(pady=(10, 10), padx=10, fill="both", expand=False)

filepathIDQ = None
filepathVOL = None

marketplace = ["PL", "SE"]

# Load IDQ report button
label = customtkinter.CTkLabel(master=frame, text="IDQ report:", font=('Roboto', 20))
label.pack(pady=(15, 5), padx=10)

def load_IDQ_and_update_prompt():
    file_loaded = load_IDQ()
    if file_loaded:
        loaded_prompt.configure(text="File loaded!")
    else:
        loaded_prompt.configure(text="")

browse_idq = customtkinter.CTkButton(master=frame, text='Browse files', command=load_IDQ_and_update_prompt)
browse_idq.pack(pady=(8, 5), padx=10)

loaded_prompt = customtkinter.CTkLabel(master=frame, text="", font=('Roboto', 12))
loaded_prompt.pack(pady=(5, 8), padx=10)

# Load Adhocvol report button
label = customtkinter.CTkLabel(master=frame, text="Adhocvol report:", font=('Roboto', 20))
label.pack(pady=(20, 5), padx=10)

def load_VOL_and_update_prompt():
    file2_loaded = load_VOL()
    if file2_loaded:
        loaded_prompt2.configure(text="File loaded!")
    else:
        loaded_prompt2.configure(text="")

browse_adhocvol = customtkinter.CTkButton(master=frame, text='Browse files', command=load_VOL_and_update_prompt)
browse_adhocvol.pack(pady=(8, 5), padx=10)

loaded_prompt2 = customtkinter.CTkLabel(master=frame, text="", font=('Roboto', 12))
loaded_prompt2.pack(pady=(5, 8), padx=10)

# Choose marketplace, year and week

marketplace_text = customtkinter.CTkLabel(master=frame, text="Marketplace, Year, Week:", font=('Roboto', 20))
marketplace_text.pack(pady=(20, 5), padx=10)
marketplace_dropdown = customtkinter.CTkOptionMenu(master=frame, values=marketplace)
marketplace_dropdown.pack(pady=10, padx=10)

year_input = customtkinter.CTkEntry(master=frame, placeholder_text="E.g. 2024")
year_input.pack(pady=0, padx=10)

week_input = customtkinter.CTkEntry(master=frame, placeholder_text="E.g. 12")
week_input.pack(pady=(10, 20), padx=10)

# weekly quality table

info_label = customtkinter.CTkLabel(
    master=frame,
    text="*Weekly quality table will be loaded automatically, based on provided year and week",
    font=('Roboto', 10),
    wraplength=240  # Wrap text at 280 pixels width
)
info_label.pack(pady=(0, 20), padx=10)



def create_report_wrapper():
    try:
        year = int(year_input.get())
        week = int(week_input.get())
        create_report(marketplace_dropdown.get(), year, week)
        create_prompt.configure(text="Report created successfully!")
    except ValueError:
        create_prompt.configure(text="Invalid input for year or week!")
    except Exception as e:
        create_prompt.configure(text=f"Error: {e}")

# Create new report
createbutton = customtkinter.CTkButton(master=button_frame, text='Create report', font=('Roboto', 12, 'bold'), fg_color='Red4', command=create_report_wrapper)
createbutton.pack(pady=20, padx=0, ipady=5, ipadx=5)

create_prompt = customtkinter.CTkLabel(master=info_frame, text="", font=('Roboto', 12))
create_prompt.pack(pady=20, padx=0, ipady=5, ipadx=5)

root.mainloop()
