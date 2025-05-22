import tkinter as tk
from gui.main import ExcelProcessorGUI

def main():
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()