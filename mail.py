import tkinter as tk
from tkinter import Label, Text, Scrollbar
import threading
import cv2
import speech_recognition as sr
import pyttsx3
import smtplib
import imaplib
import email
import re
from email.message import EmailMessage
from email.header import decode_header

# Initialize Text-to-Speech Engine
listener = sr.Recognizer()
engine = pyttsx3.init()

# Gmail Credentials (Replace with your credentials)
SENDER_EMAIL = "bhavanas2802@gmail.com"
APP_PASSWORD = "ezxwtpkqzaojxoow"

# GUI Setup
root = tk.Tk()
root.title("Voice Email Assistant")
root.geometry("500x400")

status_label = Label(root, text="Voice-Controlled Email Assistant", font=("Arial", 14), pady=10)
status_label.pack()

status_text = Text(root, height=15, wrap=tk.WORD)
status_text.pack()

scrollbar = Scrollbar(root, command=status_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
status_text.config(yscrollcommand=scrollbar.set)

def talk(text):
    """Speaks and displays the given text."""
    update_status(text)
    engine.say(text)
    engine.runAndWait()

def update_status(message):
    """Update status label in the GUI."""
    status_text.insert(tk.END, message + "\n")
    status_text.see(tk.END)
    root.update()

def get_info(timeout=10, retries=3):
    """Captures and processes speech input from the user with retries."""
    for attempt in range(retries):
        try:
            with sr.Microphone() as source:
                update_status("Listening...")
                talk('Listening...')
                listener.adjust_for_ambient_noise(source, duration=1)
                voice = listener.listen(source, timeout=timeout, phrase_time_limit=10)
                info = listener.recognize_google(voice).strip().lower()
                update_status(f"Recognized: {info}")
                return info
        except sr.UnknownValueError:
            talk("I didn't catch that. Can you please repeat?")
        except sr.RequestError:
            talk("Network issue. Please check your connection.")
        except Exception as e:
            update_status(f"Error: {e}")

    talk("I couldn't understand. Moving forward.")
    return None

def format_email(email_input):
    """Formats spoken email addresses correctly."""
    email_input = email_input.lower().replace(" ", "")
    email_input = email_input.replace("attherate", "@").replace("at the rate", "@")
    email_input = email_input.replace("dot", ".")
    email_input = re.sub(r'(\d+)([a-zA-Z])', r'\1\2', email_input)
    return email_input

def validate_email(email_address):
    """Validates email format using regex."""
    email_pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(email_pattern, email_address))

def get_valid_email():
    """Gets and validates recipient's email from voice input."""
    talk("Who do you want to send an email to? Please speak the email address.")
    while True:
        email_input = get_info()
        if email_input:
            formatted_email = format_email(email_input)
            update_status(f"Converted Email: {formatted_email}")
            if validate_email(formatted_email):
                talk(f"Recipient email is {formatted_email}")
                return formatted_email
            else:
                talk("Invalid email address. Please repeat the recipient's email.")

def send_email(receiver, subject, message):
    """Sends an email via Gmail SMTP."""
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        email_msg = EmailMessage()
        email_msg['From'] = SENDER_EMAIL
        email_msg['To'] = receiver
        email_msg['Subject'] = subject
        email_msg.set_content(message)
        server.send_message(email_msg)
        server.quit()
        talk("Email sent successfully.")
    except Exception as e:
        talk("Failed to send email. Please check your credentials and internet connection.")
        update_status(f"Error: {e}")

def get_email_info():
    """Takes recipient's email, subject, and message through voice commands."""
    receiver = get_valid_email()
    talk("What is the subject of your email?")
    subject = get_info() or "No Subject"
    talk("Tell me the text in your email.")
    message = get_info() or "No message provided."
    send_email(receiver, subject, message)
    talk("Would you like to check unread emails? say please do")
    response = get_info()
    if response and "please do" in response:
        read_emails()
    else:
        talk("Okay, exiting.")

def detect_face():
    """Uses OpenCV to detect the user's face before allowing email access."""
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    cap = cv2.VideoCapture(0)
    talk("Face recognition for authentication.")

    while True:
        ret, frame = cap.read()
        if not ret:
            continue

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
            talk("Face detected. Access granted.")
            cap.release()
            cv2.destroyAllWindows()
            return True

        cv2.imshow('Face Authentication', frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    talk("Authentication failed. Please try again.")
    return False


def read_emails():
    """Reads the latest unread emails in a human-readable format."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(SENDER_EMAIL, APP_PASSWORD)
        mail.select("inbox")
        _, search_data = mail.search(None, "UNSEEN")
        email_ids = search_data[0].split()

        if not email_ids:
            talk("No new emails.")
            return

        for email_id in email_ids:
            _, msg_data = mail.fetch(email_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    # Decode email sender
                    from_email = msg["From"]

                    # Decode subject
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")

                    # Decode email body
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            # Read only plain text content, ignore attachments
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                body_bytes = part.get_payload(decode=True)  # Decode base64/quoted-printable
                                body = body_bytes.decode("utf-8", errors="ignore")
                                break
                    else:
                        body_bytes = msg.get_payload(decode=True)
                        body = body_bytes.decode("utf-8", errors="ignore")

                    # Clean up unwanted characters or links
                    body = re.sub(r'\n+', '\n', body).strip()  # Remove excessive newlines

                    talk(f"Email from {from_email}. Subject is {subject}. Message: {body[:200]}...")  # Read first 200 chars
                    update_status(f"Email from: {from_email}\nSubject: {subject}\nMessage: {body}\n")

        mail.logout()
    except Exception as e:
        talk("Failed to read emails. Please check your internet connection.")
        update_status(f"Error: {e}")


def start_process():
    """Start the email assistant process."""
    update_status("Starting authentication...")
    if detect_face():
        update_status("Face recognized! Proceeding...")
        talk("You are authenticated. Would you like to send an email or read unread emails?")
        action = get_info()
        if action:
            if "send" in action:
                get_email_info()
            elif "read" in action:
                read_emails()
            else:
                talk("Invalid choice. Please say either 'send email' or 'read emails'.")
                start_process()
        else:
            talk("No valid command detected. Exiting.")
    else:
        update_status("Authentication failed. Exiting...")

def on_start():
    """Run the process in a separate thread."""
    threading.Thread(target=start_process).start()

start_button = tk.Button(root, text="Start", font=("Arial", 14), command=on_start)
start_button.pack()

root.mainloop()

