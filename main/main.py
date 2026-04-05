"""WBS Enhancement Application — Entry Point.

Lightweight launcher that creates the Tk root window and starts
the Application GUI. All logic lives in the app, file_parser,
api_client, and wbs_writer modules.
"""

import tkinter as tk

from app import Application


if __name__ == '__main__':
    root = tk.Tk()
    root.title("WBS Enhancement")
    app = Application(master=root)
    app.mainloop()