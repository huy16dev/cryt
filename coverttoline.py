import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def convert_vbs_lines(content):
    lines = content.strip().splitlines()
    converted = []
    for i, line in enumerate(lines):
        line = line.rstrip().replace('"', '""') 
        part = f'"{line}"'
        if i < len(lines) - 1:
            part += ' & vbCrLf &'
        converted.append(part)
    return ' '.join(converted)

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("VBScript Files", "*.vbs"), ("All Files", "*.*")])
    if not file_path:
        return

    with open(file_path, 'r', encoding='utf-8') as f:
        original = f.read()

    converted = convert_vbs_lines(original)
    input_text.delete(1.0, tk.END)
    input_text.insert(tk.END, original)
    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, converted)

def save_output():
    output = output_text.get(1.0, tk.END).strip()
    if not output:
        messagebox.showwarning("chưa có nội dung", "nothing.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("VB String", "*.txt"), ("All Files", "*.*")])
    if file_path:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(output)
        messagebox.showinfo("Đã lưu", f"Đã lưu file:\n{file_path}")

root = tk.Tk()
root.title("VBS to One-Line Converter")
root.geometry("800x600")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10, fill="both", expand=True)

btn_open = tk.Button(frame, text="mở vbs", command=open_file)
btn_open.pack(pady=5)

lbl_input = tk.Label(frame, text="gốc:")
lbl_input.pack(anchor="w")

input_text = scrolledtext.ScrolledText(frame, height=10)
input_text.pack(fill="both", expand=True, padx=5, pady=5)

lbl_output = tk.Label(frame, text="đã chuyển thành chuỗi 1 dòng:")
lbl_output.pack(anchor="w")

output_text = scrolledtext.ScrolledText(frame, height=10)
output_text.pack(fill="both", expand=True, padx=5, pady=5)

btn_save = tk.Button(frame, text="💾 Save", command=save_output)
btn_save.pack(pady=10)

root.mainloop()
