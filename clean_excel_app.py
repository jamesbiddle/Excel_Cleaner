import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os
import clean_excel as ce


class App(tk.Tk):
    def __init__(self, parent=None):
        tk.Tk.__init__(self, parent)
        # Initialise frame
        self.parent = parent
        self.winfo_toplevel().title("Excel cleaner")

        self.place_window()
        self.setup_frame()

    def setup_frame(self):
        """Setup the initial look of the app
        """
        self.grid()
        self.name = None
        self.input_file = None
        for i in range(3):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Add labels and buttons
        self.lbl_infl = tk.Label(self,
                                 text="*No file selected*",
                                 anchor="center")
        self.btn_infl = tk.Button(
            self,
            text="Select input sheet",
            anchor="center",
            command=self.open_input_file,
        )

        self.btn_run = tk.Button(self,
                                 text="Run",
                                 anchor="center",
                                 command=self.run_cleaner)

        self.btn_infl.grid(row=0)
        self.lbl_infl.grid(row=1)
        self.btn_run.grid(row=2)

    def open_input_file(self):
        """Select a document to clean.
        """
        name = fd.askopenfilename()
        self.input_file = name
        if self.input_file:
            self.lbl_infl["text"] = name
        else:
            self.lbl_infl["text"] = "*No file selected*"

    def place_window(self, width=400, height=300):
        """Set the window size, and place it in the centre of the screen

        Args:
            width (int, optional): Width of the app. Defaults to 400.
            height (int, optional): Height of the app. Defaults to 300.
        """
        w = width
        h = height

        # get screen width and height
        ws = self.winfo_screenwidth()  # width of the screen
        hs = self.winfo_screenheight()  # height of the screen

        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)

        # set the dimensions of the screen
        # and where it is placed
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def run_cleaner(self):
        """Execute the excel cleaning code
        """
        in_file = self.input_file
        if in_file is None:
            showinfo("Error", "Please select an input file first.")
            return
        out_file = fd.asksaveasfilename(defaultextension=".xlsx")
        if not out_file:
            return

        in_file, out_file = ce.set_names(in_file, out_file)

        if in_file == out_file:
            showinfo("Error", "Input and output file are the same")
            return

        df = ce.clean_sheet(in_file)

        # Output new sheet
        ce.output_sheet(df, out_file)
        in_base = os.path.basename(in_file)
        out_base = os.path.basename(out_file)
        showinfo(
            "Success!",
            f"Successfully cleaned {in_base}.\n Saved cleaned sheet {out_base}",
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
