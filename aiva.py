import os
import re
import time
import datetime
import webbrowser
import pyautogui
import keyboard
import pyttsx3
import pyjokes
import speech_recognition as sr
import urllib.parse
import win32com.client
import re

# Initialize speech recognition and text-to-speech
recognizer = sr.Recognizer()
aiva = pyttsx3.init()
aiva.setProperty('rate', 150)

# Excel variables - initialized as None until needed
excel = None
wb = None
ws = None

def initialize_excel():
    """Initialize Excel only when needed"""
    global excel, wb, ws
    try:
        if excel is None:
            speak("Opening Excel for you")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Add()
            ws = wb.ActiveSheet
            speak("Excel is now ready")
        return wb.ActiveSheet  # Always return the active sheet
    except Exception as e:
        print(f"Excel initialization error: {e}")
        speak("I had trouble opening Excel. Please check if it's installed correctly.")
        return None

def speak(text):
    print("AIVA:", text)
    aiva.say(text)
    aiva.runAndWait()

def listen():
    with sr.Microphone() as source:
        print("🎤 Listening...")
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        try:
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=7)
            command = recognizer.recognize_google(audio)
            print("You:", command)
            return command.lower()
        except sr.WaitTimeoutError:
            return ""
        except sr.UnknownValueError:
            return ""
        except sr.RequestError:
            speak("My speech service is down.")
            return ""

def handle_excel_command(command):
    command = command.lower()
    print(f"Processing Excel command: {command}")
    
    # Initialize Excel if not already initialized
    ws = initialize_excel()
    if ws is None:
        speak("I couldn't access Excel to perform that command")
        return

    try:
        if "create new sheet" in command:
            wb.Sheets.Add().Name = f"Sheet{wb.Sheets.Count + 1}"
            speak("Created a new sheet")

        elif "rename sheet to" in command:
            match = re.search(r"rename sheet to (.+)", command)
            if match:
                ws.Name = match.group(1).strip()
                speak(f"Renamed sheet to {match.group(1).strip()}")

        elif "write" in command and "in cell" in command:
            match = re.search(r"write (.+) in cell ([a-z]+[0-9]+)", command)
            if match:
                text, cell = match.groups()
                ws.Range(cell).Value = text.strip()
                speak(f"Wrote '{text.strip()}' in cell {cell}")

        elif "fill column" in command:
            match = re.search(r"fill column ([a-z]) with (.+)", command)
            if match:
                col, values = match.groups()
                values = [v.strip() for v in re.split(r"[,\n]", values)]
                for i, val in enumerate(values):
                    ws.Range(f"{col.upper()}{i+1}").Value = val
                speak(f"Filled column {col.upper()} with {len(values)} values")

        elif "sum formula in cell" in command:
            match = re.search(r"sum formula in cell ([a-z]+[0-9]+) for ([a-z]+[0-9]+) to ([a-z]+[0-9]+)", command)
            if match:
                cell, start, end = match.groups()
                ws.Range(cell).Formula = f"=SUM({start}:{end})"
                speak(f"Added sum formula in cell {cell}")

        elif "average in cell" in command:
            match = re.search(r"average in cell ([a-z]+[0-9]+) for ([a-z]+[0-9]+) to ([a-z]+[0-9]+)", command)
            if match:
                cell, start, end = match.groups()
                ws.Range(cell).Formula = f"=AVERAGE({start}:{end})"
                speak(f"Added average formula in cell {cell}")

        elif "maximum in cell" in command:
            match = re.search(r"maximum in cell ([a-z]+[0-9]+) for ([a-z]+[0-9]+) to ([a-z]+[0-9]+)", command)
            if match:
                cell, start, end = match.groups()
                ws.Range(cell).Formula = f"=MAX({start}:{end})"
                speak(f"Added maximum formula in cell {cell}")

        elif "minimum in cell" in command:
            match = re.search(r"minimum in cell ([a-z]+[0-9]+) for ([a-z]+[0-9]+) to ([a-z]+[0-9]+)", command)
            if match:
                cell, start, end = match.groups()
                ws.Range(cell).Formula = f"=MIN({start}:{end})"
                speak(f"Added minimum formula in cell {cell}")

        elif "count in cell" in command:
            match = re.search(r"count in cell ([a-z]+[0-9]+) for ([a-z]+[0-9]+) to ([a-z]+[0-9]+)", command)
            if match:
                cell, start, end = match.groups()
                ws.Range(cell).Formula = f"=COUNT({start}:{end})"
                speak(f"Added count formula in cell {cell}")

        elif "bold row" in command:
            match = re.search(r"bold row ([0-9]+)", command)
            if match:
                row = int(match.group(1))
                ws.Rows(row).Font.Bold = True
                speak(f"Made row {row} bold")

        elif "align center column" in command:
            match = re.search(r"align center column ([a-z]+)", command)
            if match:
                col = match.group(1).upper()
                ws.Columns(col).HorizontalAlignment = -4108
                speak(f"Aligned column {col} to center")

        elif "currency format column" in command:
            match = re.search(r"currency format column ([a-z]+)", command)
            if match:
                col = match.group(1).upper()
                ws.Columns(col).NumberFormat = "$#,##0.00"
                speak(f"Applied currency format to column {col}")

        elif "autofit columns" in command:
            ws.Columns.AutoFit()
            speak("Auto-fitted all columns")

        elif "freeze top row" in command:
            excel.ActiveWindow.SplitRow = 1
            excel.ActiveWindow.FreezePanes = True
            speak("Froze the top row")

        elif "freeze first column" in command:
            excel.ActiveWindow.SplitColumn = 1
            excel.ActiveWindow.FreezePanes = True
            speak("Froze the first column")

        elif "add hyperlink in cell" in command:
            match = re.search(r"add hyperlink in cell ([a-z]+[0-9]+) to (https?://\S+)", command)
            if match:
                cell, url = match.groups()
                ws.Hyperlinks.Add(Anchor=ws.Range(cell), Address=url)
                speak(f"Added hyperlink in cell {cell}")

        else:
            print("Command not recognized for Excel automation.")
            speak("I didn't understand that Excel command")
    except Exception as e:
        print(f"Excel command error: {e}")
        speak("I had trouble with that Excel command")

def extract_email_command(command):
    """Extract email details from voice command"""
    command = command.lower()
    
    # Check if this is an email command
    if not any(phrase in command for phrase in ["send a mail", "send an email", "send mail", "email to"]):
        return None
    
    # Various patterns for email commands
    patterns = [
        r"send (?:a |an )?(?:mail|email) to (.+?)(?:regarding|about|for|with subject|subject) (.+)",
        r"(?:mail|email) to (.+?)(?:regarding|about|for|with subject|subject) (.+)",
        r"send (?:a |an )?(?:mail|email) to (.+?) (?:regarding|about|for|with subject|subject) (.+)"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, command)
        if match:
            recipient = match.group(1).strip()
            subject = match.group(2).strip()
            return (recipient, subject)
    
    # Alternative parsing
    if any(phrase in command for phrase in ["send a mail to", "send an email to", "send mail to", "email to"]):
        for phrase in ["send a mail to", "send an email to", "send mail to", "email to"]:
            if phrase in command:
                rest = command.split(phrase, 1)[1].strip()
                # Look for "regarding" or similar words
                for word in ["regarding", "about", "for", "with subject", "subject"]:
                    if f" {word} " in rest:
                        recipient, subject = rest.split(f" {word} ", 1)
                        return (recipient.strip(), subject.strip())
    
    # If we find email command but couldn't parse recipient/subject, prompt for them
    if any(phrase in command for phrase in ["send a mail", "send an email", "send mail", "email"]):
        speak("Who would you like to email?")
        recipient = listen()
        if not recipient:
            speak("I couldn't hear the recipient. Let's try again.")
            return None
            
        speak(f"What should be the subject of your email to {recipient}?")
        subject = listen()
        if not subject:
            speak("I couldn't hear the subject. Let's try again.")
            return None
            
        return (recipient, subject)
    
    return None

def process_email_address(recipient):
    """Process recipient string to extract or format email address"""
    # Common email domains to help with voice recognition
    common_domains = ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "icloud.com"]
    
    # Clean up initial input - remove extraneous whitespace
    recipient = recipient.strip()
    
    # Check if it already looks like an email address
    if "@" in recipient and "." in recipient.split("@")[1]:
        # Remove any spaces that might be in the username part (before @)
        parts = recipient.split("@")
        username = parts[0].replace(" ", "")
        return f"{username}@{parts[1]}"
        
    # If they said "at" instead of @ symbol
    if " at " in recipient:
        parts = recipient.split(" at ")
        # Remove spaces from username part
        username = parts[0].replace(" ", "").strip()
        domain = parts[1].strip()
        
        # Check if domain needs a common extension
        if "." not in domain:
            for common in common_domains:
                if domain in common:
                    domain = common
                    break
            else:
                # Default to gmail if no match
                domain = f"{domain}.com"
                
        return f"{username}@{domain}"
    
    # If they say the username and domain separately with "dot"
    if " dot " in recipient:
        recipient = recipient.replace(" dot ", ".")
    
    # If no @ symbol was detected, ask for clarification
    if "@" not in recipient:
        speak(f"I heard '{recipient}'. Is this an email address or a name?")
        response = listen()
        
        if "name" in response or "yes" in response:
            speak("What is the email domain? For example, gmail.com")
            domain = listen()
            if domain:
                # Clean up domain if needed
                domain = domain.replace(" dot ", ".").strip()
                if "." not in domain:
                    domain = f"{domain}.com"
                # Remove spaces from the username
                username = recipient.replace(" ", "")
                return f"{username}@{domain}"
        elif "email" in response and "@" in response:
            # They clarified with the full email
            email_pattern = r'[\w\.-]+@[\w\.-]+'
            match = re.search(email_pattern, response)
            if match:
                return match.group(0)
    
    # Still no @ symbol, make best guess
    if "@" not in recipient:
        # Remove spaces and make a valid email address
        username = recipient.replace(" ", "").lower()
        return f"{username}@gmail.com"
    
    return recipient

def generate_email_body(subject):
    """Generate appropriate email template based on subject keywords"""
    subject_lower = subject.lower()
    
    # Sick leave template
    if any(keyword in subject_lower for keyword in ["sick", "illness", "health", "unwell", "leave"]):
        return f"""Dear Sir/Madam,

I am writing to inform you that I am unable to attend work today due to illness. I anticipate returning to work once I have recovered.

Please let me know if you need any additional information.

Thank you for your understanding.

Best regards,
AIVA User"""
    
    # Meeting template
    elif any(keyword in subject_lower for keyword in ["meeting", "appointment", "discussion", "conference"]):
        return f"""Dear Recipient,

I am writing regarding the {subject}. I would like to confirm our meeting and discuss any necessary preparations or agenda items.

Please let me know if the scheduled time works for you or if you need any changes.

Best regards,
AIVA User"""
    
    # General template
    else:
        return f"""Dear Recipient,

I am writing regarding {subject}.

Please let me know if you need any additional information.

Best regards,
AIVA User"""

def open_gmail_compose(recipient=None, subject=None):
    """Open Gmail compose with pre-filled fields if provided"""
    base_url = "https://mail.google.com/mail/?view=cm&fs=1"
    
    # Process recipient if provided
    if recipient:
        formatted_recipient = process_email_address(recipient)
        print(f"DEBUG - Formatted email address: {formatted_recipient}")
        base_url += f"&to={urllib.parse.quote(formatted_recipient)}"
    
    # Process subject if provided
    if subject:
        base_url += f"&su={urllib.parse.quote(subject)}"
    
    # Add body text if we have a subject to generate from
    if subject:
        body = generate_email_body(subject)
        encoded_body = urllib.parse.quote(body)
        base_url += f"&body={encoded_body}"
    
    print(f"DEBUG - Opening URL: {base_url}")
    
    # Open Gmail compose
    webbrowser.open(base_url)
    speak("Opening Gmail with your email draft")
    
    # Pause to let the browser open and page load
    time.sleep(3)
    
    return True

def send_email_via_keyboard():
    """Send email using keyboard shortcuts with improved reliability - FIXED VERSION"""
    try:
        # Wait longer for Gmail interface to load completely
        time.sleep(5)
        
        # Make sure the browser window is in focus
        pyautogui.click(500, 500)  # Click somewhere in the middle of the window to ensure focus
        time.sleep(1)
        
        # Use Gmail's primary keyboard shortcut to send an email
        print("Sending email with Ctrl+Enter shortcut")
        pyautogui.hotkey('ctrl', 'enter')
        time.sleep(2)
        
        speak("Email sent")
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def process_command(command):
    # Debug message
    print(f"Processing command: '{command}'")
    
    # Excel command detection and keywords - expanded to catch more phrases
    excel_keywords = ["excel", "spreadsheet", "cell", "column", "row", "formula", 
                     "new sheet", "create new sheet", "rename sheet", "write in cell", 
                     "fill column", "sum formula", "average", "maximum", "minimum", 
                     "count", "bold row", "align center", "currency format", 
                     "autofit", "freeze", "hyperlink"]
                     
    # Check if command contains any Excel keywords
    is_excel_command = any(keyword in command.lower() for keyword in excel_keywords)
    
    if is_excel_command:
        handle_excel_command(command)
        return
        
    # Try to extract email command
    email_details = extract_email_command(command)
    
    if email_details:
        recipient, subject = email_details
        speak(f"Sending an email to {recipient} regarding {subject}")
        
        # Open Gmail with pre-filled content
        opened = open_gmail_compose(recipient, subject)
        
        if opened:
            # Send email automatically without asking for confirmation
            if send_email_via_keyboard():
                speak("I've sent the email for you.")
            else:
                speak("I couldn't send the email automatically. You can send it manually when ready.")
        
        return
    
    # Handle other commands
    elif "open gmail" in command:
        webbrowser.open("https://mail.google.com")
        speak("Opening Gmail")
    
    elif "open excel" in command:
        initialize_excel()
    
    elif "time" in command:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"The current time is {current_time}")

    elif "open youtube" in command:
        webbrowser.open("https://youtube.com")
        speak("Opening YouTube")

    elif "joke" in command:
        speak(pyjokes.get_joke())

    elif "search for" in command:
        query = command.replace("search for", "").strip()
        webbrowser.open(f"https://www.google.com/search?q={urllib.parse.quote(query)}")
        speak(f"Searching for {query}")

    elif "open chrome" in command:
        try:
            # Try common Chrome paths
            chrome_paths = [
                "C:/Program Files/Google/Chrome/Application/chrome.exe",
                "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
                # Add other common paths
            ]
            
            for path in chrome_paths:
                if os.path.exists(path):
                    os.startfile(path)
                    speak("Opening Chrome")
                    time.sleep(2)
                    return
                    
            speak("I couldn't find Chrome. Please check your installation.")
        except Exception as e:
            speak("I couldn't open Chrome.")
            print("Error:", e)

    elif "new tab" in command:
        pyautogui.hotkey('ctrl', 't')
        speak("Opened a new tab")

    elif "close tab" in command:
        pyautogui.hotkey('ctrl', 'w')
        speak("Closed the current tab")

    elif "scroll down" in command:
        pyautogui.scroll(-1000)
        speak("Scrolling down")

    elif "scroll up" in command:
        pyautogui.scroll(1000)
        speak("Scrolling up")

    elif "refresh" in command:
        pyautogui.hotkey('ctrl', 'r')
        speak("Refreshing page")

    elif "open chatgpt" in command:
        webbrowser.open("https://chat.openai.com")
        speak("Opening ChatGPT")

    elif "shutdown" in command:
        speak("Are you sure you want to shut down your computer? Say yes or no.")
        confirmation = listen()
        if "yes" in confirmation:
            os.system("shutdown /s /t 1")
            speak("Shutting down")
        else:
            speak("Shutdown canceled")

    elif "restart" in command:
        speak("Are you sure you want to restart your computer? Say yes or no.")
        confirmation = listen()
        if "yes" in confirmation:
            os.system("shutdown /r /t 1")
            speak("Restarting the system")
        else:
            speak("Restart canceled")

    elif "lock screen" in command:
        os.system("rundll32.exe user32.dll,LockWorkStation")
        speak("Locking your screen")

    elif "who are you" in command:
        speak("I am AIVA, your Artificial Intelligence Virtual Assistant. I can help you with tasks like sending emails, opening applications, searching the web, and more.")

    elif "what can you do" in command:
        speak("I can help you with several tasks like sending emails through Gmail, telling time, opening applications, searching the web, controlling browser tabs, telling jokes, and basic system controls.")

    else:
        speak("Sorry, I don't understand that command.")

# Main program start
if __name__ == "__main__":
    speak("Hello, I am AIVA. How can I assist you today?")
    
    # Main loop
    while True:
        cmd = listen()
        if cmd:
            if "exit" in cmd or "stop" in cmd or "goodbye" in cmd:
                speak("Goodbye from AIVA! Have a great day!")
                break
            process_command(cmd)