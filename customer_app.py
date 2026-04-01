import tkinter as tk
from app.main_window import CustomerSystem


if __name__ == "__main__":
    root = tk.Tk()
    window_width = 1400
    window_height = 850
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    app = CustomerSystem(root)
    root.mainloop()
