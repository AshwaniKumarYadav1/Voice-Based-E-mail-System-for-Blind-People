import speech_recognition as sr
import smtplib
import imaplib
import email
import os
import time
import threading
from gtts import gTTS
import pyglet
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import tkinter as tk
from tkinter import scrolledtext

# Load environment variables
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
if not EMAIL_USER or not EMAIL_PASS:
    raise ValueError("Please set EMAIL_USER and EMAIL_PASS in a .env file.")

# Utility Functions
def speak(text, filename="temp.mp3"):
    """Convert text to speech and play it."""
    try:
        tts = gTTS(text=text, lang='en')
        tts.save(filename)
        music = pyglet.media.load(filename, streaming=False)
        music.play()
        time.sleep(music.duration)
        os.remove(filename)
    except Exception as e:
        log_to_gui(f"Error in TTS: {e}")

def log_to_gui(message):
    """Log messages to the GUI text area."""
    text_area.config(state=tk.NORMAL)
    text_area.insert(tk.END, f"{message}\n")
    text_area.see(tk.END)
    text_area.config(state=tk.DISABLED)

def get_voice_input(prompt, retries=2, timeout=3, phrase_limit=3):
    """Capture voice input with enhanced recognition, no confirmation."""
    speak(prompt)
    log_to_gui(f"Listening for '{prompt}'")
    r = sr.Recognizer()

    try:
        with sr.Microphone() as source:
            log_to_gui("Microphone detected.")
            r.adjust_for_ambient_noise(source, duration=0.5)
            log_to_gui("Ambient noise adjusted.")
    except sr.RequestError as e:
        speak("Microphone not detected. Please check your setup.")
        log_to_gui(f"ERROR: Microphone initialization failed: {e}")
        return None

    for attempt in range(retries):
        try:
            with sr.Microphone() as source:
                log_to_gui(f"{prompt} (Attempt {attempt + 1}/{retries})")
                audio = r.listen(source, timeout=timeout, phrase_time_limit=phrase_limit)
                log_to_gui("Audio captured.")
                
                result = r.recognize_google(audio).lower()
                log_to_gui(f"Recognized: '{result}'")
                return result
        except sr.UnknownValueError:
            speak("Sorry, I didn’t catch that. Please speak clearly.")
            log_to_gui("Audio not understood.")
        except sr.RequestError as e:
            speak(f"Network error. Trying again.")
            log_to_gui(f"ERROR: Google API failed: {e}")
        except sr.WaitTimeoutError:
            speak("No input detected. Speak louder or check your microphone.")
            log_to_gui("Timeout - no audio detected.")
        except Exception as e:
            speak(f"Error: {e}. Retrying.")
            log_to_gui(f"ERROR: Unexpected issue: {e}")

    speak("Failed to recognize after retries. Please try again.")
    return None

# Email Functions
def speech_compose():
    """Compose and send an email with subject, message, and automatic @gmail.com."""
    log_to_gui("Starting compose process.")
    subject = get_voice_input("Please say the subject of your email:")
    if not subject:
        return

    message = get_voice_input("Please say your message:")
    if not message:
        return

    first_part = get_voice_input("Say the first part of the recipient’s Gmail address (e.g., 'test' for 'test@gmail.com'):")
    if not first_part:
        speak("Couldn’t hear the email. Aborting.")
        return

    first_part = first_part.replace(" ", "")
    recipient = first_part + "@gmail.com"
    log_to_gui(f"Recognized recipient: '{recipient}'")

    full_message = f"Subject: {subject}\n\n{message}"
    
    try:
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()
        mail.starttls()
        log_to_gui(f"Attempting SMTP login with {EMAIL_USER}")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.sendmail(EMAIL_USER, recipient, full_message)
        speak("Congratulations! Your email has been sent.")
        log_to_gui("Email sent successfully.")
        mail.close()
    except smtplib.SMTPAuthenticationError as e:
        speak("Authentication failed. Check your email credentials.")
        log_to_gui(f"SMTP Auth Error: {e}")
    except smtplib.SMTPException as e:
        speak(f"Failed to send email: {str(e)}. Please try again.")
        log_to_gui(f"SMTP Error: {e}")

def inbox():
    """Check inbox details and read the latest email."""
    log_to_gui("Checking inbox.")
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        log_to_gui(f"Attempting IMAP login with {EMAIL_USER}")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select('Inbox')

        status, total = mail.select('Inbox')
        total_count = int(total[0].decode())
        speak(f"Total emails in your inbox: {total_count}")
        log_to_gui(f"Total emails: {total_count}")

        status, unseen = mail.search(None, 'UNSEEN')
        unseen_count = len(unseen[0].split())
        speak(f"You have {unseen_count} unread emails.")
        log_to_gui(f"Unread emails: {unseen_count}")

        status, data = mail.search(None, 'ALL')
        inbox_items = data[0].split()
        if not inbox_items:
            speak("Your inbox is empty.")
            log_to_gui("Inbox is empty.")
            mail.logout()
            return

        latest_email_id = inbox_items[-1]
        status, email_data = mail.fetch(latest_email_id, '(RFC822)')
        raw_email = email_data[0][1].decode("utf-8", errors="ignore")
        email_message = email.message_from_string(raw_email)

        sender = email_message['From']
        subject = email_message['Subject'] or "No subject"
        speak(f"Latest email from: {sender}. Subject: {subject}")
        log_to_gui(f"From: {sender}\nSubject: {subject}")

        if email_message.is_multipart():
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                    break
        else:
            body = email_message.get_payload(decode=True).decode("utf-8", errors="ignore")
        
        cleaned_body = BeautifulSoup(body, "html.parser").get_text().strip()
        speak(f"Body: {cleaned_body[:200]}")
        log_to_gui(f"Body: {cleaned_body[:200]}")

        mail.logout()
    except imaplib.IMAP4.error as e:
        speak("Authentication failed for inbox. Check your credentials.")
        log_to_gui(f"IMAP Auth Error: {e}")
    except Exception as e:
        speak(f"Error checking inbox: {str(e)}. Please try again.")
        log_to_gui(f"IMAP Error: {e}")

def search():
    """Search for unread emails from a specific username."""
    log_to_gui("Starting search process.")
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        log_to_gui(f"Attempting IMAP login with {EMAIL_USER}")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select('Inbox')

        username = get_voice_input("Say the first part of the sender’s Gmail address (e.g., 'test' for 'test@gmail.com'):")
        if not username:
            speak("Couldn’t hear the username. Aborting.")
            mail.logout()
            return

        username = username.replace(" ", "")
        search_email = username + "@gmail.com"
        log_to_gui(f"Searching for unread emails from: '{search_email}'")

        status, data = mail.search(None, 'UNSEEN', f'FROM "{search_email}"')
        email_ids = data[0].split()
        unread_count = len(email_ids)
        speak(f"You have {unread_count} unread emails from {search_email}.")
        log_to_gui(f"Found {unread_count} unread emails from {search_email}")

        if unread_count == 0:
            speak("No unread emails found from this sender.")
            mail.logout()
            return

        for email_id in email_ids:
            status, email_data = mail.fetch(email_id, '(RFC822)')
            raw_email = email_data[0][1].decode("utf-8", errors="ignore")
            email_message = email.message_from_string(raw_email)

            sender = email_message['From']
            subject = email_message['Subject'] or "No subject"
            speak(f"Email from: {sender}. Subject: {subject}")
            log_to_gui(f"From: {sender}\nSubject: {subject}")

            if email_message.is_multipart():
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        break
            else:
                body = email_message.get_payload(decode=True).decode("utf-8", errors="ignore")
            
            cleaned_body = BeautifulSoup(body, "html.parser").get_text().strip()
            speak(f"Body: {cleaned_body[:200]}")
            log_to_gui(f"Body: {cleaned_body[:200]}")

        mail.logout()
    except imaplib.IMAP4.error as e:
        speak("Authentication failed for search. Check your credentials.")
        log_to_gui(f"IMAP Auth Error: {e}")
    except Exception as e:
        speak(f"Error searching inbox: {str(e)}. Please try again.")
        log_to_gui(f"IMAP Error: {e}")

# Voice Command Loop
def voice_command_loop():
    """Continuously listen for voice commands and trigger actions."""
    options = {
        "compose": speech_compose,
        "inbox": inbox,
        "search": search,
        "quit": lambda: (speak("Goodbye."), root.quit())
    }

    while True:
        choice = get_voice_input("Say 'compose', 'inbox', 'search', or 'quit':")
        if not choice:
            continue

        action = options.get(choice)
        if action:
            action()
            if choice == "quit":
                break
        else:
            speak("Invalid choice. Please say 'compose', 'inbox', 'search', or 'quit'.")
        time.sleep(1)  # Brief pause to avoid overlap

# GUI Setup
root = tk.Tk()
root.title("Voice-Based Email System")
root.geometry("600x400")

# Text area for logs
text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=20, state=tk.DISABLED)
text_area.pack(pady=10)

# Label for instructions
label = tk.Label(root, text="Speak 'compose', 'inbox', 'search', or 'quit' to proceed.")
label.pack(pady=5)

# Initial welcome message and start voice loop
def start_app():
    speak("Welcome to the Voice-Based Email System for the Blind.")
    log_to_gui(f"Logged in as: {os.getlogin()}")
    threading.Thread(target=voice_command_loop, daemon=True).start()

root.after(100, start_app)  # Run after GUI is initialized
root.mainloop()