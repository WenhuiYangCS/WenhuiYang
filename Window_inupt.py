

"""
Final version
"""
import math
import tkinter as tk


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None

        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if not self.tooltip:
            x, y, _, _ = self.widget.bbox("insert")
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 25

            self.tooltip = tk.Toplevel(self.widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")

            label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, padx=5, pady=2)
            label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

def create_label_frames(root, frame_names, labels_list, default_values_list=None, help_texts_list=None):
    label_frames = []  # List to hold the created LabelFrame widgets
    entry_values = []  # List to hold the entry values

    for name, labels, default_values, help_texts in zip(frame_names, labels_list, default_values_list or [], help_texts_list or []):
        # Create a LabelFrame with the specified name
        label_frame = tk.LabelFrame(root, text=name, labelanchor="n", padx=10, pady=10)
        label_frame.pack(side="left", anchor="n", padx=5)  # Pack them horizontally

        # List to hold the entry values for the current LabelFrame
        frame_entry_values = []

        # Create labels, entry widgets, and help texts within the LabelFrame based on the labels_list and help_texts_list
        for i, (label_text, default_value, help_text) in enumerate(zip(labels, default_values, help_texts)):
            label = tk.Label(label_frame, text=label_text)
            label.grid(row=i, column=0, sticky="w")  # Place label in column 0, aligned to the west (left)

            entry = tk.Entry(label_frame,width=10)
            entry.insert(0, default_value)  # Set default value for the entry widget
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")  # Place entry widget in column 1, with padding and alignment to east-west

            # Add tooltip for help text
            tooltip = ToolTip(entry, help_text)

            frame_entry_values.append(entry)  # Append entry widget to the list

        label_frames.append(label_frame)  # Add the created LabelFrame to the list
        entry_values.append(frame_entry_values)  # Add entry values for the current LabelFrame to the list

    return label_frames, entry_values


def get_entry_values():
    global entry_values
    entry_values = []
    for frame_entries in entries:
        frame_values = []
        for entry in frame_entries:
            frame_values.append(float(entry.get()))
        entry_values.append(frame_values)
    print(entry_values)
    root.destroy()

# Example usage:
root = tk.Tk()
#root.iconbitmap("download.ico")
root.title("Check parameters input window _ WY")
frame_names = ["Tunnel info",
               "Joint check",
               "SLU/SLE"]
labels_list = [["D_inner [m]", "Thickness [m]", "Width [m]", "Rck [Mpa]", "fyk [Mpa]"],
               ["mdo [mm]", "mdi [mm]", "L_concio [m]"],
               ["D_lower bars [cm]", "D_upper bars [cm]", "No. Lower bars [-]", "No. Upper bars [-]", "c' [cm]", "c [cm]", "wlim [mm]"]]

default_values_list = [[8, 0.45, 1.8, 45, 450],
               [82.5, 30, 1.8],
               [1.2, 1.2, 14, 14, 6.6, 6.6, 0.2]]

help_texts_list = [["Tunnel inner diameter", "Lining thickness", "Lining width", "C35/45: 45\nC45/55: 55\nC50/60: 60", "fyk[Mpa]"],
               ["mdo [mm]", "mdi [mm]", "Lconcio [m]"],
               ["Inside bars diameter", "Outside bars diameter", "Inside bars number", "Outside bars number", "Inside concrete cover", "Outside concrete cover", "wlim [mm]"]]
frames, entries = create_label_frames(root, frame_names, labels_list, default_values_list, help_texts_list)

entry_values = []
def get_entry_values():
    global entry_values
    for frame_entries in entries:
        frame_values = []
        for entry in frame_entries:
            frame_values.append(float(entry.get()))
        entry_values.append(frame_values)
    print(entry_values)
    root.destroy()

button = tk.Button(root, text="Get Entry Values", command=get_entry_values)
button.pack()
root.mainloop()

[[D_in, L_t, L_w, R_ck, f_yk],
 [m_do, m_di, L_co],
 [D_bar_in, D_bar_out, N_bar_in, N_bar_out, c_in, c_out, wlim]] = entry_values

Af_in = (0.25 * (math.pi * D_bar_in**2)) * N_bar_in
Af_out = (0.25 * (math.pi * D_bar_out**2)) * N_bar_out
B_calc = 2 * L_w


print(f"{D_in = };", f"{L_t = };", f"{L_w = };", f"{R_ck = };", f"{f_yk = };",
 f"{m_do = };", f"{m_di = };", f"{L_co = };",
 f"{D_bar_in = };", f"{D_bar_out = };", f"{N_bar_in = };", f"{N_bar_out = };", f"{c_in = };", f"{c_out = };", f"{wlim = };", f"{Af_in = };")

import xlwings as xw

# Select folder and excel file
destination_file = r"C:\Pini\Plaxis - Python\test folder\plots to doc\Test ver 7\Joint_checks_file.xlsx"
cells_input = ["AN4","AO4", "AP4", "AS4", "BB4"]
values_input = [m_do, m_di, L_t, L_co, R_ck]
dest_wb = xw.Book(destination_file)
dest_ws = dest_wb.sheets['CONTATTO&FRETTAGGIO_DIF_INTR']
for cell_input, value_input in zip(cells_input, values_input):
    dest_ws.range(cell_input).value = value_input


# Change the parameters of "CONTATTO&FRETTAGGIO_DIF_INTR" by the input values from input window
destination_file_2 = r"C:\Pini\Plaxis - Python\test folder\plots to doc\Test ver 7\Verification_SLU_file.xlsm"
cells_input = ["D4", "D6", "D7", "D9", "D10", "D11", "D12", "D13", "D14", "D17",]
values_input = [D_bar_in, Af_in, D_bar_out, Af_out, 100 * B_calc, 100 * L_t, c_out, c_in, R_ck, f_yk]
dest_wb_2 = xw.Book(destination_file_2)
dest_ws_2 = dest_wb_2.sheets['Taglio']
for cell_input, value_input in zip(cells_input, values_input):
    dest_ws_2.range(cell_input).value = value_input


# Change the parameters of "Dati INPUT" by the input values from input window
destination_file_3 = r"C:\Pini\Plaxis - Python\test folder\plots to doc\Test ver 7\Verification_SLE_file.xls"
cells_input = ["J69", "J71", "J72", "J74", "J75", "J76", "J77", "J78", "J79", "J82",]
values_input = [D_bar_in, Af_in, D_bar_out, Af_out, 100 * B_calc, 100 * L_t, c_out, c_in, R_ck, f_yk]
dest_wb_3 = xw.Book(destination_file_3)
dest_ws_3 = dest_wb_3.sheets['Dati INPUT']
for cell_input, value_input in zip(cells_input, values_input):
    dest_ws_3.range(cell_input).value = value_input
dest_wb_3.sheets['S.L.E. Fessurazione'].range("N3").value = wlim
