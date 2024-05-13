import os
import pyaudio
import pyperclip
import pythoncom
import threading
import time
import wave
import win32com.client as win32
import customtkinter as tk
import openai
from dotenv import load_dotenv
from template import prompt_template
import logging

# Load environment variables from .env file
load_dotenv()

# Set up logging to file
logging.basicConfig(
    filename='app.log',
    filemode='w',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)

# Set up OpenAI API credentials
openai_api_key = os.getenv("OPENAI_API_KEY")
openai.api_key = openai_api_key

# Class to handle custom errors
class CustomError(Exception):
    def __init__(self, message="Custom error occurred."):
        self.message = message
        super().__init__(self.message)

# Main TK window app class
class CustomApp(tk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("EMaiL Generative Assistant")
        self.geometry("650x350")

        # Add your customtkinter widgets here
        self.label_main_1 = tk.CTkLabel(self, text="Select an email in outlook you'd like to respond to", font=("Arial", 20))
        self.label_main_1.pack(pady=5)

        self.label_main_2 = tk.CTkLabel(self, text="Record a brief response and we will create a professional response", font=("Arial", 12), wraplength=540)
        self.label_main_2.pack(pady=5)

        self.button_main = tk.CTkButton(self, text="Start Recording", command=self.start_recording)
        self.button_main.pack(pady=5)

    # Begin main thread
    def start_recording(self):
        threading.Thread(target=self.main).start()

    def get_selected_email_body_and_item(self):
        try:
            outlook = win32.Dispatch("Outlook.Application")
            explorer = outlook.ActiveExplorer()
            if explorer.Selection.Count > 0:
                item = explorer.Selection.Item(1)
                if hasattr(item, "Body"):
                    return item.Body, item
            return "No email selected or the selected item is not an email.", None
        except Exception as e:
            raise CustomError(f"Error in get_selected_email_body_and_item: {e}")

    #Record audio from the microphone and save it to a file, stopping when the button is pressed or after max_seconds
    def record_audio(self, filename="speech.wav", max_seconds=120):
        chunk = 1024
        sample_format = pyaudio.paInt16
        channels = 1
        fs = 16000
        p = pyaudio.PyAudio()
        logging.info(f"PyAudio instance created: {p}")

        try:
            stream = p.open(format=sample_format,
                            channels=channels,
                            rate=fs,
                            frames_per_buffer=chunk,
                            input=True)
            logging.info(f"Audio stream opened: {stream}")
        except Exception as e:
            logging.error(f"Error opening audio stream: {e}")
            return

        frames = []  # Initialize array to store frames
        start_time = time.time()

        # Use an event to signal when to stop recording
        stop_event = threading.Event()

        # Function to stop recording
        def stop_recording():
            stop_event.set()

        # Add your customtkinter widgets here
        self.label_rec_1 = tk.CTkLabel(self, text="Please describe the email now - be sure to state if you'd like a short, medium, or long email", wraplength=540, justify="center")
        self.label_rec_1.pack(pady=10)

        self.button_rec = tk.CTkButton(self, text="Press this button when you have finished talking", command=stop_recording)
        self.button_rec.pack(pady=0)

        # Start the recording in a separate thread
        def record():
            while not stop_event.is_set() and (time.time() - start_time < max_seconds):
                data = stream.read(chunk)
                frames.append(data)
            stream.stop_stream()
            stream.close()
            p.terminate()

        # Start recording in a new thread
        record_thread = threading.Thread(target=record)
        record_thread.start()

        # Wait for the recording to complete if still running
        if record_thread.is_alive():
            record_thread.join()

        # Save the recorded data as a WAV file
        with wave.open(filename, 'wb') as wf:
            wf.setnchannels(channels)
            wf.setsampwidth(p.get_sample_size(sample_format))
            wf.setframerate(fs)
            wf.writeframes(b''.join(frames))

    # Transcribe speech from an audio file using OpenAI Whisper
    def recognize_speech_from_whisper(self):
        filename = "speech.wav"
        self.record_audio(filename)

        with open(filename, "rb") as audio_file:
            openai.api_key = openai_api_key  # Ensure the OpenAI API key is set correctly
            try:
                # Assuming the response is plain text as per your logs
                transcription_text = openai.Audio.transcribe(
                    model="whisper-1",
                    file=audio_file,
                    response_format="text"
                )
                # Check if the transcription_text contains actual text
                if transcription_text:
                    print("Whisper recognized: " + transcription_text)
                    return {"success": True, "error": None, "transcription": transcription_text}
                else:
                    # If no transcription is detected, set it to ":)"
                    print("No transcription results.")
            except Exception as e:
                print(f"An exception occurred while processing the transcription: {e}")

            # If there is any issue with transcription or it's empty, set to ":)"
            return {"success": True, "error": "Defaulting to :)", "transcription": ":)"}


    # Generate an email using GPT based on the prompt
    def generate_email(self, prompt: str) -> str:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-0125",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=600,
            temperature=0.7
        )
        return response['choices'][0]['message']['content'].strip()

    # Create or reply to an email in Outlook with the generated email body
    def create_email_draft(self, email_body: str, selected_email_item):
        outlook = win32.Dispatch("Outlook.Application")

        if selected_email_item:
            # Check if the selected item has the ReplyAll method
            if hasattr(selected_email_item, 'ReplyAll'):
                mail = selected_email_item.ReplyAll()
                mail.Body = email_body
            else:
                # Fall back to creating a new mail item if ReplyAll is not applicable
                mail = outlook.CreateItem(0)
                mail.Body = email_body
        else:
            # Create a new email
            mail = outlook.CreateItem(0)
            mail.Body = email_body

        mail.Display(True)  # Show the email draft

        # Reset the UI (simulate 'Select new email' functionality)
        self.clear_rec()

    def clear_rec(self):
        # Remove recording labels and buttons
        try:
            self.label_rec_1.pack_forget()
            self.label_rec_2.pack_forget()
            self.button_rec.pack_forget()
            self.button_rec2.pack_forget()
            self.button_rec3.pack_forget()
        except AttributeError:
            pass  # Handle the case where some widgets were not created due to errors

        # Reset to initial state with the "Start Recording" button
        self.button_main.pack(pady=5)
        pythoncom.CoUninitialize()

    def input_speech(self):
        def submit():
            text = entry.get()

            entry.pack_forget()
            submit_button.pack_forget()
            label.pack_forget()

            self.finalise_email(text)

        # Label entry stuff
        label = tk.CTkLabel(self, text="Please enter your response manually, recording failed", font=("Arial", 12))
        label.pack(pady=10)

        # Create an Entry widget
        entry = tk.CTkEntry(self, width=400)
        entry.pack(pady=0)

        # Create a button to submit the text
        submit_button = tk.CTkButton(self, text="Submit", command=submit)
        submit_button.pack(pady=5)

    def finalise_email(self, message):
        self.label_rec_2 = tk.CTkLabel(self, text=f"You said: {message}", font=("Arial", 12))
        self.label_rec_2.pack(pady=15)

        pythoncom.CoInitialize()
        conversation_history, selected_email_item = self.get_selected_email_body_and_item()
        
        # Generate the email using GPT
        full_prompt = prompt_template.format(
            conversation_history=conversation_history,
            speech_to_text_transcription=message
        )
        email_response = self.generate_email(full_prompt)

        # Reply to the selected email or start a new email draft in Outlook
        pyperclip.copy(email_response)

        # Automatically show the email draft in Outlook and reset the UI afterward
        self.create_email_draft(email_response, selected_email_item)

    def main(self):
        try:
            # Speech-to-Text with Whisper
            speech_to_text = self.recognize_speech_from_whisper()
            
            if not speech_to_text["success"] or not speech_to_text["transcription"]:
                # Fallback to manual input if speech recognition fails or no transcription
                self.input_speech()
            else:
                message = speech_to_text["transcription"]
                self.finalise_email(message)
        except Exception as e:
            print("An exception occurred:", e)
            self.input_speech()  # Fallback to manual input on any error

if __name__ == "__main__":
    app = CustomApp()
    app.mainloop()
