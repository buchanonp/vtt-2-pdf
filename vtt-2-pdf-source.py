import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, Button
import requests
from html.parser import HTMLParser
import re
import nltk
from nltk.tokenize import sent_tokenize
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import pagesizes
from textwrap import wrap, fill
import time

# Ensure NLTK 'punkt' package is downloaded once
nltk.download('punkt', quiet=True)

class TitleParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_title_tag = False
        self.title = ""

    def handle_starttag(self, tag, attrs):
        if tag == 'title':
            self.in_title_tag = True

    def handle_endtag(self, tag):
        if tag == 'title':
            self.in_title_tag = False

    def handle_data(self, data):
        if self.in_title_tag:
            self.title += data

def get_title_from_html_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    parser = TitleParser()
    parser.feed(html_content)
    return parser.title.strip(), html_content

def clean_filename(title, max_length=50):
    filename = re.sub(r'[^\w\s-]', '', title)
    return filename[:max_length]

def extract_video_token(html_content):
    match = re.search(r"fast.wistia.net/embed/iframe/(.+?)\?", html_content)
    return match.group(1) if match else None

def download_vtt_file(vtt_url):
    try:
        response = requests.get(vtt_url)
        response.raise_for_status()
        return response.text
    except requests.ConnectionError:
        messagebox.showerror("Network Error", "Unable to connect to the server.")
        return None
    except requests.HTTPError as e:
        messagebox.showerror("HTTP Error", f"HTTP error occurred: {e}")
        return None
    except requests.RequestException as e:
        messagebox.showerror("Error", f"Error downloading VTT file: {e}")
        return None

def format_vtt_to_paragraphs(vtt_content):
    text_content = re.sub(r'\d{2}:\d{2}:\d{2}\.\d{3} --> \d{2}:\d{2}:\d{2}\.\d{3}', '', vtt_content).replace('WEBVTT', '').strip()
    text_content = re.sub(r'\s+', ' ', text_content)
    sentences = sent_tokenize(text_content)
    formatted_paragraphs = []
    paragraph = ''
    for sentence in sentences:
        if len(paragraph) + len(sentence) > 500:
            formatted_paragraphs.append(paragraph.strip())
            paragraph = sentence
        else:
            paragraph += ' ' + sentence
    if paragraph:
        formatted_paragraphs.append(paragraph.strip())
    return formatted_paragraphs

def save_to_pdf(formatted_paragraphs, pdf_path, title):
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = pagesizes.letter
    margin = 72
    line_height = 14
    current_height = height - margin - 40

    wrapped_title = fill(title, width=55)
    title_lines = wrapped_title.split('\n')

    c.setFont("Helvetica-Bold", 16)
    for line in title_lines:
        c.drawString(margin, current_height, line)
        current_height -= line_height
    current_height -= line_height

    c.setFont("Helvetica", 12)
    for paragraph in formatted_paragraphs:
        wrapped_text = wrap(paragraph, width=80)
        for line in wrapped_text:
            current_height -= line_height
            if current_height < margin:
                c.showPage()
                current_height = height - margin
            c.drawString(margin, current_height, line)

        current_height -= line_height

    c.save()

def show_instruction_popup(message, root):
    popup = Toplevel(root)
    popup.title("Instruction")
    popup.geometry("400x150")

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (400 / 2))
    y_coordinate = int((screen_height / 2) - (150 / 2))
    popup.geometry(f"+{x_coordinate}+{y_coordinate}")

    Label(popup, text=message, font=("Helvetica", 12, "bold"), wraplength=350).pack(padx=20, pady=20)
    Button(popup, text="OK", command=popup.destroy).pack(pady=(0, 10))
    root.wait_window(popup)

def delete_html_files(files_to_delete):
    for file_path in files_to_delete:
        try:
            os.remove(file_path)
        except OSError as e:
            messagebox.showerror("Error", f"Error: {e.strerror}")

def closing_message(root):
    closing_popup = Toplevel(root, borderwidth=4, relief="raised")
    closing_popup.title("Closing")
    closing_popup.configure(bg='red')
    closing_popup.geometry("300x100")

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (150))
    y_coordinate = int((screen_height / 2) - (50))
    closing_popup.geometry(f"+{x_coordinate}+{y_coordinate}")

    Label(closing_popup, text="VTT-to-PDF program is closing.", font=("Helvetica", 12), bg='red').pack(expand=True)
    root.update()
    time.sleep(2)
    closing_popup.destroy()

def generate_pdf_process(root):
    show_instruction_popup("Please select one or more HTML files to upload.", root)
    file_paths = filedialog.askopenfilenames(
        filetypes=[("HTML files", "*.html"), ("HTML files", "*.htm")],
        title="Select HTML Files"
    )

    if not file_paths:
        return

    show_instruction_popup("Please select the folder where to save the PDF files.", root)
    save_directory = filedialog.askdirectory()

    if not save_directory:
        return

    for file_path in file_paths:
        title, html_content = get_title_from_html_file(file_path)
        filename_title = clean_filename(title)

        video_token = extract_video_token(html_content)
        if not video_token:
            messagebox.showerror("Error", f"Video token not found in {file_path}.")
            continue

        vtt_url = f"https://fast.wistia.net/embed/captions/{video_token}.vtt?language=eng"
        vtt_content = download_vtt_file(vtt_url)
        if vtt_content is None:
            continue

        formatted_paragraphs = format_vtt_to_paragraphs(vtt_content)

        pdf_file_path = os.path.join(save_directory, f"{filename_title}.pdf")
        save_to_pdf(formatted_paragraphs, pdf_file_path, title)

    messagebox.showinfo("PDFs Saved", f"All formatted PDFs saved in: {save_directory}")

    if messagebox.askquestion("Continue", "Do you want to generate another PDF?") == 'yes':
        generate_pdf_process(root)
    else:
        if messagebox.askquestion("Delete Files", "Do you want to delete the original HTML files?") == 'yes':
            delete_html_files(file_paths)
        closing_message(root)
        root.after(2000, root.destroy)

def main():
    root = tk.Tk()
    root.withdraw()
    generate_pdf_process(root)

if __name__ == '__main__':
    main()
