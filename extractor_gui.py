#!/usr/bin/env python3

import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import sys

class GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MBOX Extractor (Year/Month Organizer)")

        tk.Label(root, text="MBOX File").grid(row=0, column=0)
        self.mbox = tk.Entry(root, width=50)
        self.mbox.grid(row=0, column=1)
        tk.Button(root, text="Browse", command=self.pick_mbox).grid(row=0, column=2)

        tk.Label(root, text="Output Folder").grid(row=1, column=0)
        self.out = tk.Entry(root, width=50)
        self.out.grid(row=1, column=1)
        tk.Button(root, text="Browse", command=self.pick_out).grid(row=1, column=2)

        # Options
        self.html = tk.BooleanVar()
        self.html_all = tk.BooleanVar()
        self.sender = tk.BooleanVar()

        tk.Checkbutton(root, text="HTML (attachment emails only)", variable=self.html)\
            .grid(row=2, column=1, sticky="w")

        tk.Checkbutton(root, text="HTML (ALL emails)", variable=self.html_all)\
            .grid(row=3, column=1, sticky="w")

        tk.Checkbutton(root, text="Organize by sender (inside Year/Month)", variable=self.sender)\
            .grid(row=4, column=1, sticky="w")

        tk.Label(root, text="File types").grid(row=5, column=0)
        self.types = tk.Entry(root, width=50)
        self.types.grid(row=5, column=1)

        tk.Button(root, text="Run", bg="green", fg="white", command=self.run)\
            .grid(row=6, column=1)

    def pick_mbox(self):
        f = filedialog.askopenfilename(filetypes=[("MBOX", "*.mbox")])
        self.mbox.delete(0, tk.END)
        self.mbox.insert(0, f)

    def pick_out(self):
        f = filedialog.askdirectory()
        self.out.delete(0, tk.END)
        self.out.insert(0, f)

    def run(self):
        mbox = self.mbox.get()
        out = self.out.get()

        if not os.path.exists(mbox):
            messagebox.showerror("Error", "Invalid MBOX file")
            return

        cmd = [sys.executable, "extract.py", mbox, "-o", out]

        if self.html.get():
            cmd.append("--html")

        if self.html_all.get():
            cmd.append("--html-all")

        if self.sender.get():
            cmd.append("--by-sender")

        types = self.types.get().strip()
        if types:
            cmd.append("--types")
            cmd.extend(types.split())

        try:
            subprocess.run(cmd, check=True)
            messagebox.showinfo("Done", "Extraction complete")
        except Exception as e:
            messagebox.showerror("Error", str(e))

root = tk.Tk()
GUI(root)
root.mainloop()
