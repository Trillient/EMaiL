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
import json

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
        self.title("Email Generative Assistant")
        self.geometry("650x400")
        self.configure(bg_color="#1e1e1e")

        # Add your customtkinter widgets here
        self.label_main_1 = tk.CTkLabel(self, text="Select an email in Outlook you'd like to respond to", font=("Helvetica Neue", 20, "bold"), text_color="#ffffff")
        self.label_main_1.pack(pady=10)

        self.label_main_2 = tk.CTkLabel(self, text="Record a brief response and we will create a professional response", font=("Helvetica Neue", 12), wraplength=540, text_color="#a9a9a9")
        self.label_main_2.pack(pady=5)

        self.button_main = tk.CTkButton(self, text="Start Recording", command=self.start_recording, fg_color="#007aff", hover_color="#005bb5", text_color="#ffffff", font=("Helvetica Neue", 14))
        self.button_main.pack(pady=20)

        self.button_settings = tk.CTkButton(self, text="Settings", command=self.open_settings, fg_color="#007aff", hover_color="#005bb5", text_color="#ffffff", font=("Helvetica Neue", 12))
        self.button_settings.pack(side="top", anchor="ne", padx=10, pady=10)

        # Flags to prevent multiple windows
        self.is_recording = False
        self.is_settings_open = False

        # Lock for threading
        self.lock = threading.Lock()

    def show_custom_error(self, message):
        # Create a Toplevel window
        error_window = tk.CTkToplevel()
        error_window.title("Error")
        error_window.geometry("400x200")
        error_window.configure(fg_color="#1e1e1e")

        # Error message label
        error_label = tk.CTkLabel(error_window, text=message, font=("Helvetica Neue", 14), fg_color="#1e1e1e", text_color="#ff3b30", wraplength=380)
        error_label.pack(pady=20, padx=20)

        # OK button to close the window
        ok_button = tk.CTkButton(error_window, text="OK", command=error_window.destroy, fg_color="#ffffff", text_color="#1e1e1e")
        ok_button.pack(pady=20)

        # Center the window on the screen
        error_window.update_idletasks()
        x = (error_window.winfo_screenwidth() - error_window.winfo_width()) // 2
        y = (error_window.winfo_screenheight() - error_window.winfo_height()) // 2
        error_window.geometry(f"+{x}+{y}")

        error_window.grab_set()
        error_window.mainloop()

    def open_settings(self):
        with self.lock:
            if self.is_settings_open:
                return
            self.is_settings_open = True

        self.settings_window = tk.CTkToplevel(self)
        self.settings_window.title("Settings")
        self.settings_window.geometry("500x400")
        self.settings_window.configure(bg_color="#1e1e1e")

        # Calculate the position to place the settings window
        main_window_x = self.winfo_x()
        main_window_y = self.winfo_y()
        main_window_width = self.winfo_width()
        main_window_height = self.winfo_height()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        new_x = main_window_x + main_window_width + 10  # Place it to the right of the main window
        new_y = main_window_y

        # If the new position goes out of the screen bounds, place it on top of the main window
        if new_x + 500 > screen_width:
            new_x = main_window_x
        if new_y + 400 > screen_height:
            new_y = main_window_y

        self.settings_window.geometry(f"500x400+{new_x}+{new_y}")

        name_label = tk.CTkLabel(self.settings_window, text="What is your first name?", font=("Helvetica Neue", 12), text_color="#ffffff")
        name_label.pack(pady=(20, 5))
        self.name_entry = tk.CTkEntry(self.settings_window, width=400)
        self.name_entry.pack(pady=(0, 20))

        style_label = tk.CTkLabel(self.settings_window, text="Enter Email Examples here to show us your Email style:", font=("Helvetica Neue", 12), text_color="#ffffff")
        style_label.pack(pady=(10, 5))
        self.style_text = tk.CTkTextbox(self.settings_window, height=10, width=400, fg_color="#ffffff", text_color="#000000")
        self.style_text.pack(pady=(0, 20), padx=20, fill='both', expand=True)

        model_label = tk.CTkLabel(self.settings_window, text="Select ChatGPT Model:", font=("Helvetica Neue", 12), text_color="#ffffff")
        model_label.pack(pady=(10, 5))

        self.model_var = tk.StringVar(value="gpt-3.5-turbo-0125")
        model_radio_4 = tk.CTkRadioButton(self.settings_window, text="Advanced (GPT-4)", variable=self.model_var, value="gpt-4", text_color="#ffffff", fg_color="#007aff")
        model_radio_4.pack(pady=5)
        model_radio_3 = tk.CTkRadioButton(self.settings_window, text="Simple (GPT-3.5)", variable=self.model_var, value="gpt-3.5-turbo-0125", text_color="#ffffff", fg_color="#007aff")
        model_radio_3.pack(pady=5)

        try:
            with open("user_settings.json", "r", encoding='utf-8') as file:
                settings = json.load(file)
                self.name_entry.insert(0, settings["user_name"])
                self.style_text.insert('1.0', settings["user_email_style"])
                self.model_var.set(settings.get("chatgpt_model", "gpt-3.5-turbo-0125"))
        except FileNotFoundError:
            print("No previous settings found. Starting fresh.")

        save_button = tk.CTkButton(self.settings_window, text="Save Settings", command=self.save_settings, fg_color="#007aff", hover_color="#005bb5", text_color="#ffffff")
        save_button.pack(pady=10)

        # Ensure the settings window stays on top and has focus
        self.settings_window.lift()
        self.settings_window.attributes('-topmost', True)
        self.settings_window.focus_force()
        self.settings_window.attributes('-topmost', False)

        # Use after_idle to ensure the main window does not take focus back
        self.settings_window.after_idle(self.settings_window.focus_force)

        # Bind close event to reset the flag
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_settings_close)

    def on_settings_close(self):
        with self.lock:
            self.is_settings_open = False
        self.settings_window.destroy()

    def save_settings(self):
        settings = {
            "user_name": self.name_entry.get(),
            "user_email_style": self.style_text.get('1.0', 'end-1c'),
            "chatgpt_model": self.model_var.get()
        }
        with open("user_settings.json", "w", encoding='utf-8') as file:
            json.dump(settings, file, ensure_ascii=False, indent=4)
        self.on_settings_close()

    def load_user_settings(self):
        try:
            with open("user_settings.json", "r", encoding='utf-8') as file:
                settings = json.load(file)
            return settings["user_name"], settings["user_email_style"]
        except FileNotFoundError:
            print("Settings file not found. Starting fresh.")
            return None, None
        except json.JSONDecodeError:
            print("Error decoding settings. Check file format.")
            return None, None

    def start_recording(self):
        with self.lock:
            if self.is_recording:
                return
            self.is_recording = True
        self.button_main.pack_forget()  # Hide the Start Recording button
        threading.Thread(target=self.main).start()

    def get_selected_email_body_and_item(self):
        try:
            outlook = win32.Dispatch("Outlook.Application")
            explorer = outlook.ActiveExplorer()
            if explorer.Selection.Count > 0:
                item = explorer.Selection.Item(1)
                if hasattr(item, "Body"):
                    return item.Body, item
            raise CustomError("No email selected or the selected item is not an email.")
        except Exception as e:
            raise CustomError(f"Error accessing the selected email: {e}")

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
        self.label_rec_1 = tk.CTkLabel(self, text="Please describe the email now - be sure to state if you'd like a short, medium, or long email", wraplength=540, justify="center", text_color="#ffffff")
        self.label_rec_1.pack(pady=10)

        self.button_rec = tk.CTkButton(self, text="Press this button when you have finished talking", command=stop_recording, fg_color="#007aff", hover_color="#005bb5", text_color="#ffffff", font=("Helvetica Neue", 12))
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

    def recognize_speech_from_whisper(self):
        filename = "speech.wav"
        self.record_audio(filename)

        with open(filename, "rb") as audio_file:
            openai.api_key = openai_api_key  # Ensure the OpenAI API key is set correctly
            try:
                transcription_text = openai.Audio.transcribe(
                    model="whisper-1",
                    file=audio_file,
                    response_format="text"
                )
                if transcription_text:
                    print("Whisper recognized: " + transcription_text)
                    return {"success": True, "error": None, "transcription": transcription_text}
            except Exception as e:
                print(f"An exception occurred while processing the transcription: {e}")
            return {"success": True, "error": "Defaulting to :)", "transcription": ":)"}

    def generate_email(self, prompt: str) -> str:
        try:
            with open("user_settings.json", "r", encoding='utf-8') as file:
                settings = json.load(file)
                selected_model = settings.get("chatgpt_model", "gpt-3.5-turbo-0125")
        except FileNotFoundError:
            selected_model = "gpt-3.5-turbo-0125"

        response = openai.ChatCompletion.create(
            model=selected_model,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=4000,
            temperature=0.7
        )
        return response['choices'][0]['message']['content'].strip()

    def simple_generate_email(self, prompt: str) -> str:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-0125",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=150,
            temperature=0.5
        )
        return response['choices'][0]['message']['content'].strip()

    def create_email_draft(self, email_body: str, selected_email_item):
        outlook = win32.Dispatch("Outlook.Application")

        try:
            if selected_email_item:
                if hasattr(selected_email_item, 'ReplyAll'):
                    mail = selected_email_item.ReplyAll()
                    mail.Body = email_body
                else:
                    mail = outlook.CreateItem(0)
                    mail.Body = email_body
            else:
                mail = outlook.CreateItem(0)
                mail.Body = email_body

            mail.Display(True)
        except Exception as e:
            print(f"Error creating email draft: {e}")
            self.show_custom_error("Make sure you have an email selected in the Outlook desktop app and are not actively replying to an email currently.")
            self.clear_rec()
            return

        self.clear_rec()

    def clear_rec(self):
        try:
            self.label_rec_1.pack_forget()
            self.label_rec_2.pack_forget()
            self.button_rec.pack_forget()
            self.button_rec2.pack_forget()
            self.button_rec3.pack_forget()
        except AttributeError:
            pass
        self.button_main.pack(pady=5)
        pythoncom.CoUninitialize()
        with self.lock:
            self.is_recording = False

    def finalise_email(self, message):
        user_name, user_defined_style = self.load_user_settings()
        if user_name is None or user_defined_style is None:
            print("Error loading settings. Using default values.")
            user_name = "Default User"
            user_defined_style = "Please specify your email style in settings."

        self.label_rec_2 = tk.CTkLabel(self, text=f"{user_name} said: {message}", font=("Helvetica Neue", 12), text_color="#a9a9a9")
        self.label_rec_2.pack(pady=15)

        try:
            conversation_history, selected_email_item = self.get_selected_email_body_and_item()
        except CustomError as e:
            print(e)
            self.clear_rec()
            self.show_custom_error("Make sure you have an email selected in the Outlook desktop app and are not actively replying to an email currently.")
            return

        full_prompt = prompt_template.format(
            user_name=user_name,
            user_defined_style=user_defined_style,
            conversation_history=conversation_history,
            speech_to_text_transcription=message
        )

        email_response_1 = self.generate_email(full_prompt)
        simple_prompt = f"Based on {user_name}'s spoken input, convert their spoken words into an Email. This is what they said:\n\n'{message}'"
        email_response_2 = self.simple_generate_email(simple_prompt)

        final_email = f"{email_response_1}\n\n==\n\n{email_response_2}"
        
        pyperclip.copy(final_email)
        self.create_email_draft(final_email, selected_email_item)

    def main(self):
        try:
            speech_to_text = self.recognize_speech_from_whisper()
            
            if not speech_to_text["success"] or not speech_to_text["transcription"]:
                raise CustomError("Speech recognition failed.")
            else:
                message = speech_to_text["transcription"]
                self.finalise_email(message)
        except Exception as e:
            print("An exception occurred:", e)
            self.clear_rec()
            self.show_custom_error(str(e))

if __name__ == "__main__":
    app = CustomApp()
    app.mainloop()
