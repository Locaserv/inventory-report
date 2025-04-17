import customtkinter as ctk

app = ctk.CTk()

app.title("Generate Inventory Report - v0.1")
app.geometry("800x600")

button = ctk.CTkButton(app, text="Generate Report", command=lambda: print("Report generated!"))
button.pack(pady=20)

app.mainloop()