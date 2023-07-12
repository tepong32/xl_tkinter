import tkinter as tk
from tkinter import ttk

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")


frame = ttk.Frame(root) # parent widget
frame.pack()    # this function call makes the app resposive. Since this is
                # the root widget, adjusting the size of the UI/app will keep
                # the components centered

### col#1, row1 of the frame/root widget
widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10) # padding on x & y axis


################## "Insert Row" LabelFrame ##################
### All ttk widgets here should have "widgets_frame" as their parent
### since they are supposed to be "grouped-together".

### col1, row1 ### Entry Field
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Name")    # Placeholder: insert str("Name") at index 0
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end')) # clears the text of the placeholder from index 0 to end
name_entry.grid(row=0, column=0,padx=5, pady=(0,5), sticky="ew") # "ew" means stretch from east-west

### col1, row2 ### Spinbox (but not spinning)
age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100) # set min&max values
age_spinbox.insert(0, "Age")
# no need to delete the placeholder as with the Entry widget, it automatically deletes the value once the spinbox arrow keys are used
age_spinbox.grid(row=1, column=0, padx=5, pady=(0,5),sticky="ew")

### col1, row3 ### Dropdown
# set values for the options using a variable
combo_list = ["Subscribed", "Not Subscribed", "Other"]
# set the widget
status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0) # default value selected from combo_list var
status_combobox.grid(row=2, column=0, padx=5, pady=(0,5), sticky="ew")

### col1, row4 ### Checkbox
# set the value for the variable "cb" to be used on the checkbox
cb = tk.BooleanVar()
checkbox = ttk.Checkbutton(widgets_frame, text="Employed", variable=cb)
checkbox.grid(row=3, column=0, padx=5, pady=(0,5), sticky="nsew")

### col1, row5 ###
button = ttk.Button(widgets_frame, text="Insert")
button.grid(row=4, column=0, sticky="nsew")

### separator ###
separator = ttk.Separator(widgets_frame)
separator.grid(row=6, column=0, padx=20, pady=10, sticky="ew")

### switch (dark/light)
def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch",
    command=toggle_mode) # this triggers the toggle_mode function above
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

################## /"Insert Row" LabelFrame ##################



################## Excel LabelFrame ####################################
### This is where the preview of the excel file's data will be displayed

### Outer Frame
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Name", "Age", "Subscription","Employment") # column of the preview related to the excel file
treeView = ttk.Treeview(treeFrame, show="headings", 
                        yscrollcommand=treeScroll.set, columns=cols, height=15)
# these set the width of the columns specifically
treeView.column("Name", width=100)
treeView.column("Age", width=50)
treeView.column("Subscription", width=100)
treeView.column("Employment", width=100)
treeView.pack()
treeScroll.config(command=treeView.yview) # this line attaches the treeScroll widget to the treeView, scrolling vertically


### attaching the excel file to the UI starts here:
# import openpyxl

# def load_data():


# load_data()


root.mainloop()
