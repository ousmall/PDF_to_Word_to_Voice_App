from tkinter import *
from tkinter.ttk import Combobox
import ttkbootstrap as ttk
from tkinter import messagebox, filedialog
from google.cloud import texttospeech
from pypdf import PdfReader  # both are ok
# import pdfplumber
from docx import Document
import os

# set up Google Cloud credentials environment variables
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "path/to/your/google/cloud/credentials.json"


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        if not file_path.lower().endswith('.pdf'):
            messagebox.showerror("Error", "Please select a PDF file.")
            return
        path_textbox.delete(0, END)
        path_textbox.insert(0, file_path)


def upload_pdf():
    file_path = path_textbox.get()
    if file_path:
        text = convert_pdf_to_text(file_path)
        if text:
            text_box.delete(1.0, END)
            path_textbox.delete(0, END)
            text_box.insert(END, text)
        else:
            messagebox.showerror("Error", "Failed to convert PDF to text.")


def convert_pdf_to_text(pdf_file):
    text = ""
    with open(pdf_file, 'rb') as file:
        reader = PdfReader(file)
        num_pages = len(reader.pages)
        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text += page.extract_text()
    return text

    # text = ""
    # with pdfplumber.open(pdf_file) as pdf:
    #     for page in pdf.pages:
    #         text += page.extract_text()
    # return text


def save_as_word(text, file_path):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(file_path)


def download_word():
    text = text_box.get(1.0, END)
    default_file_name = "converted_file.docx"
    file_path = filedialog.asksaveasfilename(
        initialfile=default_file_name,
        defaultextension=".docx",
        filetypes=[("Word files", "*.docx")])
    if file_path:
        save_as_word(text, file_path)
        messagebox.showinfo("Success", f"Your file has been saved in {file_path}")


def text_to_speech():
    text = {"content": text_box.get(1.0, END)}
    language_name = language_opt.get()
    language_code = languages[language_name]

    client = texttospeech.TextToSpeechClient()
    synthesis_input = texttospeech.SynthesisInput(text=text)
    voice = texttospeech.VoiceSelectionParams(language_code=language_code)
    voice.ssml_gender = texttospeech.SsmlVoiceGender.NEUTRAL
    audio_config = texttospeech.AudioConfig()
    audio_config.audio_encoding = texttospeech.AudioEncoding.MP3

    # Call the API and generate speech
    response = client.synthesize_speech(
        input=synthesis_input, voice=voice, audio_config=audio_config
    )
    with open("output.mp3", "wb") as out:
        out.write(response.audio_content)

    # Play the generated voice file
    os.system("start output.mp3")


# --------------- GUI Setting ----------------
root = ttk.Window()
style = ttk.Style("pulse")
root.config(padx=50, pady=50)
root.title("PDF to Word to Voice Converter")

frame = Frame(root)
frame.grid(column=1, row=1, rowspan=6)
frame.config(padx=20)

notification = Label(root, text="File should clear, no special signs, unencrypted.",
                     font=("Helvetica", 10, "italic"))
notification.config(foreground="red")
notification.grid(column=1, row=1, pady=10, sticky=W, columnspan=2)

path_textbox = Entry(root, width=62)
path_textbox.grid(column=1, row=2, columnspan=2)

select_button = Button(root, text='Select PDF', command=select_file)
select_button.grid(column=1, row=3, pady=20, sticky=W)

upload_button = Button(root, text="Upload", command=upload_pdf)
upload_button.grid(column=2, row=3, sticky=W)

text_box = Text(root, height=20, width=60)
text_box.grid(column=1, row=4, columnspan=2, pady=20)

languages = {
    "English": "en-US",
    "French": "fr-FR",
    "Japanese": "ja-JP",
    "Chinese": "zh-CN"
}
language_opt = StringVar(root)
language_menu = Combobox(root, textvariable=language_opt, values=list(languages.keys()), state="readonly", width=10)
language_opt.set("English")
language_menu.grid(column=1, row=5)

text_to_speech_button = Button(root, text="Text to Speech", command=text_to_speech)
text_to_speech_button.grid(column=2, row=5)

download_button = Button(root, text="Download Word", command=download_word)
download_button.grid(column=1, row=6, columnspan=2, pady=20)


root.mainloop()
