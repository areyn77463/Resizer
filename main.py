import os
import cv2
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.chart import (
    PieChart,
    Reference,
    BarChart
)
from openpyxl.chart.label import DataLabelList
from idlelib.tooltip import Hovertip

root = tk.Tk()

# Window setup
root.title("Photo Resizer")
root.iconbitmap("Resize_Icon.ico")
root.resizable(False, False)

window_width = 400
window_height = 200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int(screen_width/2 - window_width/2)
center_y = int(screen_height/2 - window_height/2)
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

canvas1 = tk.Canvas(root, width=400, height=200)
canvas1.pack()

scale_label = tk.Label(root, text="Scale to").place(x=15, y=130)
scale_entry = tk.Entry(root, width=5)
canvas1.create_window(85, 140, window=scale_entry)
scale_label = tk.Label(root, text="% of the original dimensions").place(x=105, y=130)

# Supported file types to check
global TYPES
TYPES = [
    '.bmp',
    ',dib',
    ',jpeg',
    '.jpg',
    '.jpe',
    '.jp2',
    '.png',
    '.webp',
    '.pbm',
    '.pgm',
    '.ppm',
    '.pxm',
    '.pnm',
    '.sr',
    '.ras',
    '.tiff',
    '.tif',
    '.exr',
    '.hdr',
    '.pic'
]

# xlsx skeleton
global COLUMNS
global COL_WIDTH
global HEADERS
COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
COL_WIDTH = [15, 18, 17, 17, 16, 16, 10]
HEADERS = [
            "Original Filename",
            "Resized Filename",
            "Original Dimensions",
            "Resized Dimensions",
            "Original File Size",
            "Resized FIle Size",
            "Status"
]

global DATA
DATA = []
global PIE_DATA
PIE_DATA = [0, 0]
global BAR_DATA
BAR_DATA = [0, 0]


def check_image(filename):
    """
    Checks if file is supported
    :param filename: name of file
    :type: filename: string
    :rtype: bool
    :return: True for files with supported extension, otherwise False
    """
    if filename[-4:] in TYPES:
        return True
    return False


image_path = tk.StringVar()
from_text = tk.StringVar()
from_text.set("Image path")
image_path_label = tk.Label(root, textvariable=from_text, wraplength=340, justify='left').place(x=50, y=51)

save_path = tk.StringVar()
save_text = tk.StringVar()
save_text.set("Save destination")
save_path_level = tk.Label(root, textvariable=save_text, wraplength=340, justify='left').place(x=50, y=91)

file_count = tk.IntVar(root, 0)
resized_count = tk.IntVar(root, 0)
count_label = tk.Label(root, textvariable=resized_count).place(x=15, y=15)
between_label = tk.Label(root, text="of").place(x=30, y=15)
resize_label = tk.Label(root, textvariable=file_count).place(x=45, y=15)
after_label = tk.Label(root, text="have been resized").place(x=60, y=15)

status = tk.StringVar()
status.set("Status")
status_label = tk.Label(root, textvariable=status).place(x=270, y=15)


def get_image_path():
    """
    Sets path to original images
    :rtype: void
    :return:
    """
    DATA.clear()
    PIE_DATA[0] = 0
    PIE_DATA[1] = 0
    BAR_DATA[0] = 0
    BAR_DATA[1] = 0
    file_count.set(0)
    resized_count.set(0)
    root.update()
    image_path_var = filedialog.askdirectory()
    if image_path_var != "":
        if button3["state"] == "normal":
            button3["state"] = "disabled"
        image_path.set(image_path_var)
        from_text.set(image_path_var)
        if save_text.get() == "Save destination":
            save_path.set(image_path_var)
            save_text.set(image_path_var)
        for filename in os.listdir(image_path.get()):
            file = os.path.join(image_path.get(), filename)
            if os.path.isfile(file):
                if check_image(file):
                    file_count.set(file_count.get()+1)

    if file_count.get() > 0:
        button2["state"] = "normal"
        button1["state"] = "normal"
    else:
        button2["state"] = "disabled"
        button1["state"] = "disabled"


def get_save_path():
    """
    Sets path to save destination
    :rtype: void
    :return:
    """
    DATA.clear()
    PIE_DATA[0] = 0
    PIE_DATA[1] = 0
    BAR_DATA[0] = 0
    BAR_DATA[1] = 0
    save_path_var = filedialog.askdirectory()
    if save_path_var != "":
        if button3["state"] == "normal":
            button3["state"] = "disabled"
        save_path.set(save_path_var)
        save_text.set(save_path_var)


def calculate_file_size(size):
    """
    Converts bytes to the appropriate size
    :param size: file size on disk in bytes
    :type: size: int
    :rtype: string
    :return: converted file size
    """
    if size/1048576 > 1:
        return str(round(size/1048576, 1)) + ' MB'
    if size/1024 > 1:
        return str(round(size/1024, 2)) + ' KB'
    return str(size) + ' bytes'


def work():
    if not scale_entry.get().isnumeric() or int(scale_entry.get()) < 1 or int(scale_entry.get()) > 200:
        tk.messagebox.showerror(title="Work Alert", message="Invalid scale\n"
                                                            "1. Scale cannot be empty\n"
                                                            "2. Scale must be an integer n such that 200 ≤ n > 0")
        scale_entry.delete(0, last=tk.END)
        return

    status.set("Running...")
    root.update()
    scale = int(scale_entry.get())

    for filename in os.listdir(image_path.get()):
        file = os.path.join(image_path.get(), filename)
        if os.path.isfile(file):
            if check_image(file):
                img = cv2.imread(file, -1)

                # print("Original Dimensions: ", img.shape)
                # print("Original File Size: ", os.path.getsize(f"{image_path.get()}/{filename}"))

                width = int(img.shape[1] * (scale/100))
                height = int(img.shape[0] * (scale/100))
                dimensions = (width, height)

                resized = cv2.resize(img, dimensions)

                # print("Resize Dimensions: ", resized.shape)

                if cv2.imwrite(f"{save_path.get()}/resized_{filename}", resized):
                    # print("Resize File size: ", os.path.getsize(f"{save_path.get()}/resized_{filename}"))
                    DATA.append({
                        "originalFilename": filename,
                        "resizedFilename": f"resized_{filename}",
                        "originalDimensions": f"{str(img.shape[0])} x {str(img.shape[1])}",
                        "resizedDimensions": f"{str(resized.shape[0])} x {str(resized.shape[1])}",
                        "originalSize": calculate_file_size(os.path.getsize(f"{image_path.get()}/{filename}")),
                        "resizedSize": calculate_file_size(os.path.getsize(f"{save_path.get()}/resized_{filename}")),
                        "status": "Success"
                    })
                    PIE_DATA[0] += 1
                    BAR_DATA[0] += int(os.path.getsize(f"{image_path.get()}/{filename}"))
                    BAR_DATA[1] += int(os.path.getsize(f"{save_path.get()}/resized_{filename}"))
                    resized_count.set(resized_count.get()+1)
                    root.update()
                    # print(resized_count.get())
                else:
                    DATA.append({
                        "originalFilename": filename,
                        "resizedFilename": "",
                        "originalDimensions": f"{str(img.shape[0])} x {str(img.shape[1])}",
                        "resizedDimensions": "",
                        "originalSize": calculate_file_size(os.path.getsize(f"{image_path.get()}/{filename}")),
                        "resizedSize": "",
                        "status": "Failure"
                    })
                    PIE_DATA[1] += 1
                    BAR_DATA[0] += int(os.path.getsize(f"{image_path.get()}/{filename}"))

    status.set("Done")
    button3["state"] = "normal"
    button1["state"] = "disabled"


def preview():
    if not scale_entry.get().isnumeric() or int(scale_entry.get()) < 1 or int(scale_entry.get()) > 200:
        tk.messagebox.showerror(title="Preview Alert", message="Invalid scale\n"
                                                            "1. Scale cannot be empty\n"
                                                            "2. Scale must be an integer n such that 200 ≤ n > 0")
        scale_entry.delete(0, last=tk.END)
        return

    scale = int(scale_entry.get())
    check_file = ""
    for filename in os.listdir(image_path.get()):
        file = os.path.join(image_path.get(), filename)
        if os.path.isfile(file):
            if check_image(file):
                check_file = filename
                break

    check_file = os.path.join(image_path.get(), check_file)
    img = cv2.imread(check_file, -1)
    width = int(img.shape[1] * (scale / 100))
    height = int(img.shape[0] * (scale / 100))
    dimensions = (width, height)

    resized = cv2.resize(img, dimensions)
    cv2.imshow("Original", img)
    cv2.moveWindow("Original", 40, 30)
    cv2.imshow("Preview", resized)
    cv2.moveWindow("Preview", 40, 30)
    cv2.waitKey(0)
    cv2.destroyAllWindows()

# Was using until I found out there is messagebox in tkinter
# def open_popup(title_param, text_param):
#     top = tk.Toplevel(root)
#     top.iconbitmap("Resize_Icon.ico")
#     top.geometry("250x100")
#     top.title(title_param)
#     tk.Label(top, text=text_param).pack(expand=True)


def create_excel():

    try:
        file = os.path.join(save_path.get(), "Results.xlsx")
        if os.path.isfile(file):
            os.remove(file)
    except IOError:
        tk.messagebox.showerror(title="Details Alert", message="The action cannot be completed because the file is open"
                                                               "\n\n Close and try again")
        return

    wb = Workbook()
    sheet = wb.active

    for index, col in enumerate(COLUMNS):
        sheet.column_dimensions[col].width = COL_WIDTH[index]

    sheet.append(HEADERS)
    for data in DATA:
        sheet.append([
            data["originalFilename"],
            data["resizedFilename"],
            data["originalDimensions"],
            data["resizedDimensions"],
            data["originalSize"],
            data["resizedSize"],
            data["status"]
        ])

    sheet['H2'] = f"Success: {PIE_DATA[0]}"
    sheet['H3'] = f"Failure: {PIE_DATA[1]}"
    sheet['I2'] = PIE_DATA[0]
    sheet['I3'] = PIE_DATA[1]

    pie_chart = PieChart()
    labels = Reference(sheet, min_col=8, min_row=2, max_row=3)
    data = Reference(sheet, min_col=9, min_row=1, max_row=3)
    pie_chart.add_data(data, titles_from_data=True)
    pie_chart.title = "Status Overview"
    pie_chart.set_categories(labels)

    pie_chart.dataLabels = DataLabelList()
    pie_chart.dataLabels.showPercent = True

    sheet.add_chart(pie_chart, "H1")

    sheet['K1'] = str(BAR_DATA[0]) + " bytes"
    sheet['L1'] = str(BAR_DATA[1]) + " bytes"
    sheet['J2'] = "Original Size"
    sheet['J3'] = "Resized Size"
    sheet['K2'] = int(BAR_DATA[0])
    sheet['L3'] = int(BAR_DATA[1])

    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.style = 15
    bar_chart.title = "Size On Disk Difference"
    bar_chart.y_axis.title = "Size (bytes)"
    bar_chart.x_axis.title = "Group"

    labels = Reference(sheet, min_col=10, min_row=2, max_row=3, max_col=10)
    data = Reference(sheet, min_col=11, min_row=1, max_row=3, max_col=12)
    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(labels)
    sheet.add_chart(bar_chart, "H17")

    wb.save(f"{save_path.get()}/Results.xlsx")
    # tk.messagebox.showinfo(title="Details Alert", message="Results.xlsx successfully created")
    if messagebox.askyesno("Details Alert", "Results.xlsx successfully created\n\nWould you like to open it?"):
        os.startfile(f"{save_path.get()}/Results.xlsx")
    button3["state"] = "disabled"


def return_to_default():
    button1["state"] = "disabled"
    button2["state"] = "disabled"
    button3["state"] = "disabled"
    status.set("Status")
    save_text.set("Save destination")
    save_path.set("")
    from_text.set("Image path")
    image_path.set("")
    scale_entry.delete(0, last=tk.END)
    file_count.set(0)
    resized_count.set(0)
    DATA.clear()
    PIE_DATA[0] = 0
    PIE_DATA[1] = 0
    BAR_DATA[0] = 0
    BAR_DATA[1] = 0


dir_icon = tk.PhotoImage(file="3767084.png")
dir_icon = dir_icon.subsample(25, 25)
image_path_button = tk.Button(image=dir_icon, command=get_image_path)
canvas1.create_window(30, 60, window=image_path_button)
save_path_button = tk.Button(image=dir_icon, command=get_save_path)
canvas1.create_window(30, 100, window=save_path_button)

return_icon = tk.PhotoImage(file="return.png")
return_icon = return_icon.subsample(40, 40)
return_button = tk.Button(image=return_icon, command=return_to_default)
canvas1.create_window(350, 25, window=return_button)


image_tip = Hovertip(image_path_button, 'Set path to original\nimages')
save_tip = Hovertip(save_path_button, 'Set path to save\ndestination')
return_tip = Hovertip(return_button, 'Reset to default\noptions')
scale_tip = Hovertip(scale_entry, "Enter integer between 0 and 201\nfor the image to be resized to")


button1 = tk.Button(text='Work', command=work)
button1["state"] = "disabled"
canvas1.create_window(200, 180, window=button1)
work_tip = Hovertip(button1, "Begin resizing images")

button2 = tk.Button(text='Preview', command=preview)
button2["state"] = "disabled"
canvas1.create_window(100, 180, window=button2)
preview_tip = Hovertip(button2, "Get a resized preview\nof the original image")

button3 = tk.Button(text='Details', command=create_excel)
button3["state"] = "disabled"
canvas1.create_window(300, 180, window=button3)
details_tip = Hovertip(button3, "Generate information\nfrom the procedure")


def on_closing():
    if messagebox.askyesno("Quit", "Do you want to quit?"):
        root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()


