import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pyttsx3
import PyPDF2
import os
import pygame
import threading
import http.server
import socketserver
import webbrowser
from pydub import AudioSegment
from PIL import Image, ImageTk

selected_pdf = ""
recent_files = []
audio_text = ""
temp_audio_path = "temp_audio.wav"
is_paused = False
audio_length = 0
updating_slider = False
webserver_thread = None
port = 8000
current_voice_id = None
current_rate = 200
current_volume = 1.0
dark_mode = False

LIGHT_THEME = {
    "bg": "#f8fafc",
    "fg": "#1a202c",
    "accent": "#2b6cb0",
    "accent2": "#38b2ac",
    "surface": "#fff",
    "border": "#cbd5e1",
    "btn_bg": "#2563eb",
    "btn_fg": "#fff",
    "btn2_bg": "#38b2ac",
    "btn2_fg": "#fff",
    "input_bg": "#fff",
    "input_fg": "#1a202c",
    "select_bg": "#e3e8ee",
    "list_bg": "#f1f5f9"
}
DARK_THEME = {
    "bg": "#23272f",
    "fg": "#f8fafc",
    "accent": "#60a5fa",
    "accent2": "#38b2ac",
    "surface": "#353a45",
    "border": "#2d333b",
    "btn_bg": "#2563eb",
    "btn_fg": "#fff",
    "btn2_bg": "#38b2ac",
    "btn2_fg": "#fff",
    "input_bg": "#23272f",
    "input_fg": "#f8fafc",
    "select_bg": "#353a45",
    "list_bg": "#23272f"
}
theme = LIGHT_THEME

def update_theme():
    global theme
    theme = DARK_THEME if dark_mode else LIGHT_THEME
    window.config(bg=theme["bg"])
    for widget in window.winfo_children():
        set_widget_theme(widget)
    style.theme_use('clam')
    style.configure('TProgressbar', foreground=theme["accent"], background=theme["accent"])
    style.configure('TScale', background=theme["bg"])
    style.configure('TCombobox', fieldbackground=theme["input_bg"], background=theme["input_bg"], foreground=theme["input_fg"])
    style.configure('TListbox', background=theme["list_bg"], foreground=theme["fg"])
    style.map("TButton",
        background=[("active", theme["accent2"]), ("!active", theme["btn_bg"])],
        foreground=[("active", "#fff"), ("!active", "#fff")]
    )
    try:
        seek_slider.configure(style="TScale")
        progress_bar.configure(style='TProgressbar')
    except:
        pass

def set_widget_theme(w):
    klass = w.__class__.__name__
    try:
        if klass in ["Frame", "LabelFrame"]:
            w.config(bg=theme["bg"])
        elif klass == "Label":
            w.config(bg=theme["bg"], fg=theme["fg"])
        elif klass == "Entry":
            w.config(bg=theme["input_bg"], fg=theme["input_fg"], insertbackground=theme["input_fg"])
        elif klass == "Text" or klass == "ScrolledText":
            w.config(bg=theme["surface"], fg=theme["fg"], insertbackground=theme["fg"])
        elif klass == "Button":
            if "Save" in w.cget("text") or "Export" in w.cget("text"):
                w.config(bg=theme["btn2_bg"], fg=theme["btn2_fg"], activebackground=theme["accent2"])
            else:
                w.config(bg=theme["btn_bg"], fg=theme["btn_fg"], activebackground=theme["accent"])
        elif klass == "Listbox":
            w.config(bg=theme["list_bg"], fg=theme["fg"], selectbackground=theme["select_bg"])
        elif klass == "Scale":
            w.config(bg=theme["bg"], fg=theme["accent"])
        elif klass == "Combobox":
            w.config(background=theme["input_bg"], foreground=theme["input_fg"])
    except:
        pass
    for child in w.winfo_children():
        set_widget_theme(child)

def toggle_dark_mode():
    global dark_mode
    dark_mode = not dark_mode
    update_theme()

def browse_file():
    global selected_pdf, recent_files
    file_path = filedialog.askopenfilename(title="Select a PDF File", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        selected_pdf = file_path
        if file_path not in recent_files:
            recent_files.insert(0, file_path)
            if len(recent_files) > 5:
                recent_files.pop()
        pdf_path_entry.delete(0, tk.END)
        pdf_path_entry.insert(0, file_path)
        link_label.config(text="")
        update_history()
        preview_pdf_text()
        preview_pdf_page()

def update_history():
    history_list.delete(0, tk.END)
    for f in recent_files:
        history_list.insert(tk.END, f)

def select_from_history(event):
    sel = history_list.curselection()
    if sel:
        path = history_list.get(sel[0])
        pdf_path_entry.delete(0, tk.END)
        pdf_path_entry.insert(0, path)
        preview_pdf_text()
        preview_pdf_page()

def start_webserver():
    directory = os.path.dirname(selected_pdf)
    os.chdir(directory)
    handler = http.server.SimpleHTTPRequestHandler
    with socketserver.TCPServer(("", port), handler) as httpd:
        print(f"Serving at port {port}")
        httpd.serve_forever()

def generate_link():
    global webserver_thread
    if not selected_pdf:
        messagebox.showwarning("No File Selected", "Please browse a PDF first.")
        return
    if not webserver_thread or not webserver_thread.is_alive():
        webserver_thread = threading.Thread(target=start_webserver, daemon=True)
        webserver_thread.start()
    filename = os.path.basename(selected_pdf)
    link = f"http://127.0.0.1:{port}/{filename}"
    link_label.config(text=link, fg=theme["accent"], cursor="hand2")
    link_label.bind("<Button-1>", lambda e: webbrowser.open(link))
    return link

def preview_pdf_text(page_range=None):
    path = pdf_path_entry.get()
    if not path.endswith('.pdf'):
        text_preview.delete(1.0, tk.END)
        return
    try:
        with open(path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            pages = pdf_reader.pages
            if page_range:
                start, end = page_range
                pages = pages[start-1:end]
            for page in pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        text_preview.delete(1.0, tk.END)
        text_preview.insert(tk.END, text.strip())
    except Exception as e:
        text_preview.delete(1.0, tk.END)
        text_preview.insert(tk.END, "Error reading PDF: " + str(e))

def preview_pdf_page():
    try:
        from pdf2image import convert_from_path
        path = pdf_path_entry.get()
        if path and os.path.exists(path):
            images = convert_from_path(path, first_page=1, last_page=1, fmt="ppm")
            img = images[0]
            img.thumbnail((150, 200))
            tk_img = ImageTk.PhotoImage(img)
            pdf_image_label.config(image=tk_img)
            pdf_image_label.image = tk_img
    except Exception:
        pdf_image_label.config(image='', text="No Preview")

def convert_thread():
    threading.Thread(target=convert_pdf_to_audio).start()

def convert_pdf_to_audio():
    global audio_text
    path = pdf_path_entry.get()
    if not path.endswith('.pdf'):
        messagebox.showerror("Invalid File", "Please select a valid PDF file.")
        return
    try:
        progress_bar.start(10)
        loading_label.config(text="Converting... Please wait", fg=theme["accent2"])
        page_range = get_page_range()
        with open(path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            full_text = ""
            pages = pdf_reader.pages
            if page_range:
                start, end = page_range
                pages = pages[start-1:end]
            for page in pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
        if not full_text.strip():
            progress_bar.stop()
            loading_label.config(text="", fg=theme["fg"])
            messagebox.showerror("No Text Found", "The PDF contains no extractable text.")
            return
        text_preview.delete(1.0, tk.END)
        text_preview.insert(tk.END, full_text.strip())
        audio_text = text_preview.get(1.0, tk.END).strip()
        engine = pyttsx3.init()
        set_voice_and_rate(engine)
        engine.save_to_file(audio_text, temp_audio_path)
        engine.runAndWait()
        progress_bar.stop()
        loading_label.config(text="", fg=theme["fg"])
        update_audio_length_label()
        messagebox.showinfo("Success", "Conversion Successful! Use playback controls or save audio.")
    except Exception as e:
        progress_bar.stop()
        loading_label.config(text="", fg=theme["fg"])
        messagebox.showerror("Error", f"An error occurred:\n{e}")

def get_page_range():
    val = page_range_entry.get().strip()
    if val:
        try:
            parts = val.split('-')
            start = int(parts[0])
            end = int(parts[1]) if len(parts) > 1 else start
            return (start, end)
        except:
            messagebox.showwarning("Invalid Page Range", "Please enter page range as 'start-end' or a single page number.")
            return None
    return None

def set_voice_and_rate(engine):
    global current_voice_id, current_rate, current_volume
    voices = engine.getProperty('voices')
    if voice_combo.get() == "Male":
        for v in voices:
            if "male" in v.name.lower() or "male" in v.id.lower():
                engine.setProperty('voice', v.id)
                break
    elif voice_combo.get() == "Female":
        for v in voices:
            if "female" in v.name.lower() or "female" in v.id.lower():
                engine.setProperty('voice', v.id)
                break
    elif voice_combo.get() != "Default":
        for v in voices:
            if voice_combo.get() in v.name or voice_combo.get() in v.id:
                engine.setProperty('voice', v.id)
                break
    current_rate = rate_slider.get()
    current_volume = volume_slider.get() / 100
    engine.setProperty('rate', current_rate)
    engine.setProperty('volume', current_volume)

def play_audio():
    global audio_length
    if os.path.exists(temp_audio_path):
        pygame.mixer.init()
        pygame.mixer.music.load(temp_audio_path)
        pygame.mixer.music.set_volume(volume_slider.get() / 100)
        pygame.mixer.music.play()
        audio_length = get_audio_length(temp_audio_path)
        seek_slider.config(to=int(audio_length))
        update_audio_length_label()
        update_seek_slider()
        update_time_labels()
    else:
        messagebox.showwarning("No Audio", "Convert a PDF first.")

def stop_audio():
    if pygame.mixer.get_init():
        pygame.mixer.music.stop()

def save_audio():
    if not audio_text:
        messagebox.showwarning("No Audio", "Please convert a PDF first.")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".wav",
        filetypes=[("WAV files", "*.wav"), ("MP3 files", "*.mp3"), ("OGG files", "*.ogg")],
        title="Save Audio File"
    )
    if save_path:
        ext = os.path.splitext(save_path)[1].lower()
        if ext == ".wav":
            engine = pyttsx3.init()
            set_voice_and_rate(engine)
            engine.save_to_file(audio_text, save_path)
            engine.runAndWait()
        else:
            temp_wav = save_path.replace(ext, ".wav")
            engine = pyttsx3.init()
            set_voice_and_rate(engine)
            engine.save_to_file(audio_text, temp_wav)
            engine.runAndWait()
            sound = AudioSegment.from_wav(temp_wav)
            if ext == ".mp3":
                sound.export(save_path, format="mp3")
            elif ext == ".ogg":
                sound.export(save_path, format="ogg")
            os.remove(temp_wav)
        messagebox.showinfo("Saved", f"Audio saved to:\n{save_path}")

def export_text():
    if not audio_text:
        messagebox.showwarning("No Text", "No text to export.")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
        title="Save Extracted Text"
    )
    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(audio_text)
        messagebox.showinfo("Saved", f"Text saved to:\n{save_path}")

def get_audio_length(file_path):
    audio = AudioSegment.from_wav(file_path)
    return audio.duration_seconds

def update_seek_slider():
    global updating_slider
    if pygame.mixer.get_init() and pygame.mixer.music.get_busy():
        updating_slider = True
        current_pos = pygame.mixer.music.get_pos() / 1000
        seek_slider.set(current_pos)
        update_time_labels()
        window.after(500, update_seek_slider)
        updating_slider = False

def seek_audio(event):
    if not updating_slider and pygame.mixer.get_init():
        pygame.mixer.music.stop()
        pygame.mixer.music.play()
        messagebox.showinfo("Seek Not Supported", "Seeking is not supported for WAV files.\nSave as OGG/MP3 for seeking.")

def update_time_labels():
    curr = seek_slider.get()
    tot = seek_slider.cget('to')
    elapsed = "%02d:%02d" % (int(curr)//60, int(curr)%60)
    remaining = "%02d:%02d" % (int(tot-curr)//60, int(tot-curr)%60)
    elapsed_label.config(text=f"Elapsed: {elapsed}", fg=theme["accent2"])
    remaining_label.config(text=f"Remaining: {remaining}", fg=theme["accent"])

def update_audio_length_label():
    global audio_length
    if os.path.exists(temp_audio_path):
        audio_length = get_audio_length(temp_audio_path)
        mins = int(audio_length // 60)
        secs = int(audio_length % 60)
        audio_length_label.config(
            text=f"Audio Length: {mins:02d}:{secs:02d}", 
            fg=theme["accent"], 
            bg=theme["bg"], 
            font=("Segoe UI", 10, "bold")
        )
    else:
        audio_length_label.config(text="Audio Length: 00:00", fg=theme["accent"], bg=theme["bg"])

def on_voice_change(event=None):
    pass

def show_about():
    messagebox.showinfo("About", "PDF to Audio Converter\nBy GamerAyu25\n\nFeatures:\n- PDF to audio (WAV, MP3, OGG)\n- Web link generator for PDF\n- Text preview and editing\n- Page range selection\n- Voice/rate/volume selection\n- History & dark mode\n- Keyboard shortcuts\n- Drag and drop")

def show_help():
    help_text = (
        "How to Use:\n"
        "\n1. Browse and select a PDF file.\n"
        "2. Optionally, set page range (e.g., 2-5).\n"
        "3. Preview or edit extracted text.\n"
        "4. Select voice, rate, and volume.\n"
        "5. Click 'Convert PDF to Audio'.\n"
        "6. Use playback controls or save audio in preferred format.\n"
        "7. Generate a local web link to share the PDF on your network.\n"
        "\nExtras:\n- Drag and drop a PDF onto the window to open.\n- Keyboard shortcuts: Ctrl+O (Open), Ctrl+S (Save Audio), Ctrl+E (Export Text), Ctrl+Q (Quit), Ctrl+D (Dark Mode).\n"
    )
    messagebox.showinfo("Help", help_text)

def on_drag_drop(event):
    files = window.tk.splitlist(event.data)
    if files:
        pdf_path_entry.delete(0, tk.END)
        pdf_path_entry.insert(0, files[0])
        preview_pdf_text()
        preview_pdf_page()

def on_key(event):
    if event.state & 0x4:
        if event.keysym == 'o':
            browse_file()
        elif event.keysym == 's':
            save_audio()
        elif event.keysym == 'e':
            export_text()
        elif event.keysym == 'q':
            window.destroy()
        elif event.keysym == 'd':
            toggle_dark_mode()

window = tk.Tk()
window.title("PDF to Audio Converter + PDF Web Link")
window.geometry("860x1000")
window.config(bg=theme["bg"])

style = ttk.Style(window)
style.theme_use('clam')
style.configure('TProgressbar', foreground=theme["accent"], background=theme["accent"])
style.configure('TScale', background=theme["bg"])

menubar = tk.Menu(window, bg=theme["surface"], fg=theme["fg"])
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Open PDF (Ctrl+O)", command=browse_file)
file_menu.add_command(label="Save Audio (Ctrl+S)", command=save_audio)
file_menu.add_command(label="Export Text (Ctrl+E)", command=export_text)
file_menu.add_separator()
file_menu.add_command(label="Quit (Ctrl+Q)", command=window.destroy)
menubar.add_cascade(label="File", menu=file_menu)
settings_menu = tk.Menu(menubar, tearoff=0)
settings_menu.add_command(label="Toggle Dark Mode (Ctrl+D)", command=toggle_dark_mode)
menubar.add_cascade(label="Settings", menu=settings_menu)
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Help", command=show_help)
help_menu.add_command(label="About", command=show_about)
menubar.add_cascade(label="Help", menu=help_menu)
window.config(menu=menubar)

tk.Label(window, text="PDF to Audio Converter and PDF Web Link", font=("Arial Rounded MT Bold", 22, "bold"), bg=theme["bg"], fg=theme["accent"]).pack(pady=22)

file_frame = tk.Frame(window, bg=theme["bg"])
file_frame.pack(pady=5)
pdf_path_entry = tk.Entry(file_frame, width=45, font=("Segoe UI", 13), bd=2, relief="solid", fg=theme["input_fg"], bg=theme["input_bg"], highlightbackground=theme["border"], highlightcolor=theme["accent"])
pdf_path_entry.grid(row=0, column=0, padx=10)
tk.Button(file_frame, text="Browse PDF", command=browse_file, bg=theme["btn_bg"], fg=theme["btn_fg"], font=("Segoe UI", 12, "bold"), relief="flat", width=13).grid(row=0, column=1)
tk.Label(file_frame, text="Page Range (e.g. 2-5):", bg=theme["bg"], fg=theme["accent2"], font=("Segoe UI", 12)).grid(row=0, column=2)
page_range_entry = tk.Entry(file_frame, width=8, font=("Segoe UI", 12), bd=2, relief="solid", bg=theme["input_bg"], fg=theme["input_fg"])
page_range_entry.grid(row=0, column=3, padx=5)

history_frame = tk.Frame(window, bg=theme["bg"])
history_frame.pack(pady=5)
tk.Label(history_frame, text="Recent Files:", font=("Segoe UI", 12, "bold"), bg=theme["bg"], fg=theme["accent"]).pack(anchor="w")
history_list = tk.Listbox(history_frame, height=3, width=74, font=("Segoe UI", 10), bg=theme["list_bg"], fg=theme["fg"], selectbackground=theme["select_bg"])
history_list.pack()
history_list.bind("<<ListboxSelect>>", select_from_history)

try:
    import tkinterdnd2 as tkdnd
    dnd = tkdnd.TkinterDnD.Tk()
    window.drop_target_register(tkdnd.DND_FILES)
    window.dnd_bind('<<Drop>>', on_drag_drop)
except Exception:
    pass

pdf_image_label = tk.Label(window, bg=theme["bg"])
pdf_image_label.pack(pady=5)

tk.Button(window, text="Generate Web Link for PDF", command=generate_link, bg=theme["btn2_bg"], fg=theme["btn2_fg"], font=("Segoe UI", 12, "bold"), relief="flat", width=28).pack(pady=10)
link_label = tk.Label(window, text="", font=("Segoe UI", 12, "underline"), bg=theme["bg"])
link_label.pack(pady=5)

tk.Label(window, text="Extracted Text Preview (editable before conversion):", font=("Segoe UI", 12, "bold"), bg=theme["bg"], fg=theme["accent2"]).pack(pady=5)
text_preview = scrolledtext.ScrolledText(window, height=8, width=90, font=("Segoe UI", 11), wrap="word", bg=theme["surface"], fg=theme["fg"], insertbackground=theme["fg"], bd=1)
text_preview.pack(pady=5)

tk.Button(window, text="Export Extracted Text", command=export_text, bg=theme["btn2_bg"], fg=theme["btn2_fg"], font=("Segoe UI", 12, "bold"), relief="flat", width=22).pack(pady=5)

voice_frame = tk.Frame(window, bg=theme["bg"])
voice_frame.pack(pady=10)
engine = pyttsx3.init()
voices = engine.getProperty('voices')
voice_names = ["Default", "Male", "Female"] + [v.name for v in voices]
tk.Label(voice_frame, text="Voice:", bg=theme["bg"], fg=theme["fg"], font=("Segoe UI", 12)).grid(row=0, column=0)
voice_combo = ttk.Combobox(voice_frame, values=voice_names, width=15)
voice_combo.current(0)
voice_combo.grid(row=0, column=1, padx=5)
voice_combo.bind("<<ComboboxSelected>>", on_voice_change)
tk.Label(voice_frame, text="Rate:", bg=theme["bg"], fg=theme["fg"], font=("Segoe UI", 12)).grid(row=0, column=2)
rate_slider = tk.Scale(voice_frame, from_=100, to=300, orient="horizontal", width=10, bg=theme["bg"], fg=theme["accent"], highlightbackground=theme["border"])
rate_slider.set(200)
rate_slider.grid(row=0, column=3, padx=5)
tk.Label(voice_frame, text="Volume:", bg=theme["bg"], fg=theme["fg"], font=("Segoe UI", 12)).grid(row=0, column=4)
volume_slider = tk.Scale(voice_frame, from_=0, to=100, orient="horizontal", width=10, bg=theme["bg"], fg=theme["accent"], highlightbackground=theme["border"])
volume_slider.set(100)
volume_slider.grid(row=0, column=5, padx=5)

tk.Button(window, text="Convert PDF to Audio", command=convert_thread, bg=theme["btn_bg"], fg=theme["btn_fg"], font=("Segoe UI", 13, "bold"), relief="flat", width=22).pack(pady=15)
loading_label = tk.Label(window, text="", font=("Segoe UI", 12), bg=theme["bg"], fg=theme["accent2"])
loading_label.pack(pady=5)
progress_bar = ttk.Progressbar(window, orient="horizontal", mode="indeterminate", length=440)
progress_bar.pack(pady=5)

tk.Label(window, text="Audio Playback Controls", font=("Segoe UI", 15, "bold"), bg=theme["bg"], fg=theme["accent"]).pack(pady=13)
controls_frame = tk.Frame(window, bg=theme["bg"])
controls_frame.pack(pady=10)

play_btn = tk.Button(
    controls_frame, text="Play", command=play_audio, 
    bg=theme["btn2_bg"], fg=theme["btn2_fg"], 
    font=("Segoe UI", 12, "bold"), relief="flat", width=10
)
play_btn.grid(row=0, column=0, padx=5)

stop_btn = tk.Button(
    controls_frame, text="Stop", command=stop_audio, 
    bg="#ef4444", fg="#fff", 
    font=("Segoe UI", 12, "bold"), relief="flat", width=10
)
stop_btn.grid(row=0, column=1, padx=5)

seek_slider = ttk.Scale(window, from_=0, to=100, orient="horizontal", length=450)
seek_slider.pack(pady=10)
seek_slider.bind("<ButtonRelease-1>", seek_audio)

audio_length_label = tk.Label(window, text="Audio Length: 00:00", font=("Segoe UI", 10, "bold"), bg=theme["bg"], fg=theme["accent"])
audio_length_label.pack()

elapsed_label = tk.Label(window, text="Elapsed: 00:00", font=("Segoe UI", 10), bg=theme["bg"], fg=theme["accent2"])
elapsed_label.pack()
remaining_label = tk.Label(window, text="Remaining: 00:00", font=("Segoe UI", 10), bg=theme["bg"], fg=theme["accent"])
remaining_label.pack()

tk.Button(window, text="Save Audio as WAV/MP3/OGG", command=save_audio, bg=theme["btn2_bg"], fg=theme["btn2_fg"], font=("Segoe UI", 12, "bold"), relief="flat", width=25).pack(pady=15)

window.bind("<Control-o>", on_key)
window.bind("<Control-s>", on_key)
window.bind("<Control-e>", on_key)
window.bind("<Control-q>", on_key)
window.bind("<Control-d>", on_key)

update_theme()
window.mainloop()
