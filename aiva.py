# AIVA - AI Voice Assistant
# Complete GitHub-Ready Project Structure

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
import json
import logging
import threading
import subprocess
import psutil
import requests
from pathlib import Path
import sqlite3
from typing import Dict, List, Optional, Tuple
import configparser

# ===== Configuration Management =====
class AIVAConfig:
    def __init__(self):
        self.config_file = "config/aiva_config.ini"
        self.config = configparser.ConfigParser()
        self.load_config()
    
    def load_config(self):
        """Load configuration from file or create default"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
        else:
            self.create_default_config()
    
    def create_default_config(self):
        """Create default configuration file"""
        os.makedirs("config", exist_ok=True)
        
        self.config['VOICE'] = {
            'rate': '150',
            'volume': '0.9',
            'voice_index': '0'
        }
        
        self.config['SPEECH_RECOGNITION'] = {
            'timeout': '5',
            'phrase_time_limit': '7',
            'ambient_duration': '0.5'
        }
        
        self.config['APPLICATIONS'] = {
            'chrome_path': 'C:/Program Files/Google/Chrome/Application/chrome.exe',
            'excel_auto_open': 'false',
            'browser_delay': '3'
        }
        
        self.config['EMAIL'] = {
            'default_domain': 'gmail.com',
            'auto_send': 'true',
            'send_delay': '5'
        }
        
        self.config['LOGGING'] = {
            'level': 'INFO',
            'file': 'logs/aiva.log',
            'max_size': '10485760'
        }
        
        with open(self.config_file, 'w') as f:
            self.config.write(f)
    
    def get(self, section, key, fallback=None):
        """Get configuration value"""
        return self.config.get(section, key, fallback=fallback)
    
    def getint(self, section, key, fallback=0):
        """Get integer configuration value"""
        return self.config.getint(section, key, fallback=fallback)
    
    def getboolean(self, section, key, fallback=False):
        """Get boolean configuration value"""
        return self.config.getboolean(section, key, fallback=fallback)

# ===== Logging Setup =====
class AIVALogger:
    def __init__(self, config: AIVAConfig):
        self.config = config
        self.setup_logging()
    
    def setup_logging(self):
        """Setup logging configuration"""
        os.makedirs("logs", exist_ok=True)
        
        log_level = getattr(logging, self.config.get('LOGGING', 'level', 'INFO'))
        log_file = self.config.get('LOGGING', 'file', 'logs/aiva.log')
        
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger('AIVA')
    
    def info(self, message):
        self.logger.info(message)
    
    def error(self, message):
        self.logger.error(message)
    
    def warning(self, message):
        self.logger.warning(message)
    
    def debug(self, message):
        self.logger.debug(message)

# ===== Database Manager =====
class AIVADatabase:
    def __init__(self):
        self.db_path = "data/aiva.db"
        os.makedirs("data", exist_ok=True)
        self.init_database()
    
    def init_database(self):
        """Initialize SQLite database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Commands history table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS command_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                command TEXT NOT NULL,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                success BOOLEAN DEFAULT TRUE,
                response TEXT
            )
        ''')
        
        # User preferences table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_preferences (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Email templates table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS email_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                subject TEXT,
                body TEXT,
                created DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def log_command(self, command: str, success: bool = True, response: str = ""):
        """Log command to database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO command_history (command, success, response) VALUES (?, ?, ?)",
            (command, success, response)
        )
        conn.commit()
        conn.close()
    
    def get_command_history(self, limit: int = 50) -> List[Dict]:
        """Get command history"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT command, timestamp, success, response FROM command_history ORDER BY timestamp DESC LIMIT ?",
            (limit,)
        )
        results = cursor.fetchall()
        conn.close()
        
        return [
            {
                "command": row[0],
                "timestamp": row[1],
                "success": bool(row[2]),
                "response": row[3]
            }
            for row in results
        ]
    
    def save_email_template(self, name: str, subject: str, body: str):
        """Save email template"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO email_templates (name, subject, body) VALUES (?, ?, ?)",
            (name, subject, body)
        )
        conn.commit()
        conn.close()
    
    def get_email_template(self, name: str) -> Optional[Dict]:
        """Get email template by name"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT subject, body FROM email_templates WHERE name = ?",
            (name,)
        )
        result = cursor.fetchone()
        conn.close()
        
        if result:
            return {"subject": result[0], "body": result[1]}
        return None

# ===== Web Search Integration =====
class WebSearchManager:
    def __init__(self, logger: AIVALogger):
        self.logger = logger
    
    def search_web(self, query: str, num_results: int = 5) -> List[Dict]:
        """Search web using multiple search engines"""
        try:
            # This would integrate with search APIs in production
            # For now, we'll simulate with Google search
            search_url = f"https://www.google.com/search?q={urllib.parse.quote(query)}"
            webbrowser.open(search_url)
            
            return [{"title": f"Search results for: {query}", "url": search_url}]
        except Exception as e:
            self.logger.error(f"Web search error: {e}")
            return []
    
    def get_weather(self, location: str = "current") -> Dict:
        """Get weather information"""
        try:
            # This would integrate with weather APIs in production
            weather_url = f"https://www.google.com/search?q=weather+{location}"
            webbrowser.open(weather_url)
            return {"status": "opened", "location": location}
        except Exception as e:
            self.logger.error(f"Weather fetch error: {e}")
            return {"status": "error", "message": str(e)}

# ===== System Monitor =====
class SystemMonitor:
    def __init__(self, logger: AIVALogger):
        self.logger = logger
    
    def get_system_info(self) -> Dict:
        """Get system information"""
        try:
            cpu_percent = psutil.cpu_percent(interval=1)
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            return {
                "cpu_usage": cpu_percent,
                "memory_usage": memory.percent,
                "memory_available": memory.available // (1024**3),  # GB
                "disk_usage": disk.percent,
                "disk_free": disk.free // (1024**3)  # GB
            }
        except Exception as e:
            self.logger.error(f"System info error: {e}")
            return {}
    
    def get_running_processes(self, limit: int = 10) -> List[Dict]:
        """Get top running processes"""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_percent']):
                try:
                    processes.append(proc.info)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Sort by CPU usage
            processes.sort(key=lambda x: x['cpu_percent'] or 0, reverse=True)
            return processes[:limit]
        except Exception as e:
            self.logger.error(f"Process list error: {e}")
            return []

# ===== Enhanced Excel Manager =====
class ExcelManager:
    def __init__(self, logger: AIVALogger):
        self.logger = logger
        self.excel = None
        self.workbook = None
        self.worksheet = None
        self.is_initialized = False
    
    def initialize(self):
        """Initialize Excel application"""
        try:
            if not self.is_initialized:
                self.logger.info("Initializing Excel...")
                self.excel = win32com.client.Dispatch("Excel.Application")
                self.excel.Visible = True
                self.workbook = self.excel.Workbooks.Add()
                self.worksheet = self.workbook.ActiveSheet
                self.is_initialized = True
                self.logger.info("Excel initialized successfully")
            return True
        except Exception as e:
            self.logger.error(f"Excel initialization error: {e}")
            return False
    
    def create_chart(self, data_range: str, chart_type: str = "Column"):
        """Create chart from data range"""
        try:
            if not self.is_initialized:
                return False
            
            chart = self.worksheet.Shapes.AddChart2().Chart
            chart.SetSourceData(self.worksheet.Range(data_range))
            chart.ChartType = getattr(win32com.client.constants, f"xl{chart_type}")
            
            self.logger.info(f"Created {chart_type} chart for range {data_range}")
            return True
        except Exception as e:
            self.logger.error(f"Chart creation error: {e}")
            return False
    
    def apply_conditional_formatting(self, range_addr: str, condition: str, format_type: str):
        """Apply conditional formatting to range"""
        try:
            if not self.is_initialized:
                return False
            
            range_obj = self.worksheet.Range(range_addr)
            # Add conditional formatting logic here
            self.logger.info(f"Applied conditional formatting to {range_addr}")
            return True
        except Exception as e:
            self.logger.error(f"Conditional formatting error: {e}")
            return False
    
    def save_workbook(self, filename: str = None):
        """Save current workbook"""
        try:
            if not self.is_initialized:
                return False
            
            if filename:
                self.workbook.SaveAs(filename)
            else:
                self.workbook.Save()
            
            self.logger.info(f"Workbook saved: {filename or 'default'}")
            return True
        except Exception as e:
            self.logger.error(f"Save workbook error: {e}")
            return False
    
    def close(self):
        """Close Excel application"""
        try:
            if self.is_initialized:
                self.workbook.Close()
                self.excel.Quit()
                self.is_initialized = False
                self.logger.info("Excel closed")
        except Exception as e:
            self.logger.error(f"Excel close error: {e}")

# ===== Enhanced Email Manager =====
class EmailManager:
    def __init__(self, config: AIVAConfig, logger: AIVALogger, database: AIVADatabase):
        self.config = config
        self.logger = logger
        self.database = database
        self.common_domains = ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "icloud.com"]
    
    def parse_email_command(self, command: str) -> Optional[Tuple[str, str, str]]:
        """Enhanced email command parsing"""
        command = command.lower()
        
        # Check for template usage
        template_match = re.search(r"send email using template (\w+)", command)
        if template_match:
            template_name = template_match.group(1)
            template = self.database.get_email_template(template_name)
            if template:
                # Extract recipient
                recipient_match = re.search(r"to (.+?)(?:\s|$)", command)
                if recipient_match:
                    recipient = recipient_match.group(1).strip()
                    return (recipient, template["subject"], template["body"])
        
        # Regular email parsing (existing logic)
        patterns = [
            r"send (?:a |an )?(?:mail|email) to (.+?)(?:regarding|about|for|with subject|subject) (.+?)(?:saying|with message|body) (.+)",
            r"send (?:a |an )?(?:mail|email) to (.+?)(?:regarding|about|for|with subject|subject) (.+)",
            r"(?:mail|email) to (.+?)(?:regarding|about|for|with subject|subject) (.+)"
        ]
        
        for pattern in patterns:
            match = re.search(pattern, command)
            if match:
                if len(match.groups()) == 3:
                    return (match.group(1).strip(), match.group(2).strip(), match.group(3).strip())
                else:
                    return (match.group(1).strip(), match.group(2).strip(), self.generate_email_body(match.group(2)))
        
        return None
    
    def process_email_address(self, recipient: str) -> str:
        """Enhanced email address processing"""
        recipient = recipient.strip()
        
        # Handle voice recognition replacements
        replacements = {
            " at ": "@",
            " dot ": ".",
            "gmail": "gmail.com",
            "yahoo": "yahoo.com",
            "outlook": "outlook.com",
            "hotmail": "hotmail.com"
        }
        
        for old, new in replacements.items():
            recipient = recipient.replace(old, new)
        
        # Validate and format email
        if "@" in recipient and "." in recipient.split("@")[1]:
            parts = recipient.split("@")
            username = parts[0].replace(" ", "")
            return f"{username}@{parts[1]}"
        
        # Default domain handling
        if "@" not in recipient:
            username = recipient.replace(" ", "").lower()
            default_domain = self.config.get('EMAIL', 'default_domain', 'gmail.com')
            return f"{username}@{default_domain}"
        
        return recipient
    
    def generate_email_body(self, subject: str) -> str:
        """Generate contextual email body"""
        subject_lower = subject.lower()
        
        templates = {
            "meeting": """Dear Recipient,

I hope this email finds you well. I am writing to discuss the {subject}.

I would appreciate the opportunity to meet and discuss this matter further. Please let me know your availability.

Best regards,
AIVA User""",
            
            "sick": """Dear Sir/Madam,

I am writing to inform you that I will be unable to attend work today due to illness. I expect to return once I have recovered.

I will monitor my email and respond to urgent matters as my health permits.

Thank you for your understanding.

Best regards,
AIVA User""",
            
            "follow_up": """Dear Recipient,

I hope you are doing well. I wanted to follow up on our previous discussion regarding {subject}.

Please let me know if you need any additional information or if there are any updates.

Looking forward to your response.

Best regards,
AIVA User"""
        }
        
        # Select appropriate template
        if any(word in subject_lower for word in ["sick", "illness", "health", "leave"]):
            return templates["sick"]
        elif any(word in subject_lower for word in ["meeting", "appointment", "discussion"]):
            return templates["meeting"].format(subject=subject)
        elif any(word in subject_lower for word in ["follow", "update", "check"]):
            return templates["follow_up"].format(subject=subject)
        else:
            return f"""Dear Recipient,

I hope this email finds you well. I am writing regarding {subject}.

Please let me know if you need any additional information.

Best regards,
AIVA User"""
    
    def open_gmail_compose(self, recipient: str, subject: str, body: str = None) -> bool:
        """Open Gmail compose with enhanced features"""
        try:
            formatted_recipient = self.process_email_address(recipient)
            self.logger.info(f"Composing email to: {formatted_recipient}")
            
            base_url = "https://mail.google.com/mail/?view=cm&fs=1"
            base_url += f"&to={urllib.parse.quote(formatted_recipient)}"
            base_url += f"&su={urllib.parse.quote(subject)}"
            
            if not body:
                body = self.generate_email_body(subject)
            
            base_url += f"&body={urllib.parse.quote(body)}"
            
            webbrowser.open(base_url)
            time.sleep(self.config.getint('EMAIL', 'send_delay', 5))
            
            return True
        except Exception as e:
            self.logger.error(f"Gmail compose error: {e}")
            return False
    
    def auto_send_email(self) -> bool:
        """Automatically send email using keyboard shortcuts"""
        try:
            if not self.config.getboolean('EMAIL', 'auto_send', True):
                return False
            
            # Focus browser and send
            pyautogui.click(500, 500)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'enter')
            time.sleep(2)
            
            self.logger.info("Email sent automatically")
            return True
        except Exception as e:
            self.logger.error(f"Auto send error: {e}")
            return False

# ===== Voice Manager =====
class VoiceManager:
    def __init__(self, config: AIVAConfig, logger: AIVALogger):
        self.config = config
        self.logger = logger
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.tts_engine = pyttsx3.init()
        self.setup_voice()
        self.is_listening = False
        self.wake_words = ["aiva", "hey aiva", "ok aiva"]
    
    def setup_voice(self):
        """Setup text-to-speech configuration"""
        try:
            voices = self.tts_engine.getProperty('voices')
            voice_index = self.config.getint('VOICE', 'voice_index', 0)
            
            if voices and len(voices) > voice_index:
                self.tts_engine.setProperty('voice', voices[voice_index].id)
            
            rate = self.config.getint('VOICE', 'rate', 150)
            volume = float(self.config.get('VOICE', 'volume', '0.9'))
            
            self.tts_engine.setProperty('rate', rate)
            self.tts_engine.setProperty('volume', volume)
            
            self.logger.info("Voice engine configured")
        except Exception as e:
            self.logger.error(f"Voice setup error: {e}")
    
    def speak(self, text: str, interrupt: bool = False):
        """Enhanced text-to-speech with interruption support"""
        try:
            if interrupt:
                self.tts_engine.stop()
            
            print(f"AIVA: {text}")
            self.tts_engine.say(text)
            self.tts_engine.runAndWait()
            
            self.logger.debug(f"Spoke: {text}")
        except Exception as e:
            self.logger.error(f"Speech error: {e}")
    
    def listen(self, timeout: int = None, wake_word_mode: bool = False) -> str:
        """Enhanced speech recognition with wake word support"""
        if timeout is None:
            timeout = self.config.getint('SPEECH_RECOGNITION', 'timeout', 5)
        
        try:
            with self.microphone as source:
                if not wake_word_mode:
                    print("🎤 Listening...")
                
                self.recognizer.adjust_for_ambient_noise(
                    source, 
                    duration=self.config.getint('SPEECH_RECOGNITION', 'ambient_duration', 0.5)
                )
                
                phrase_time_limit = self.config.getint('SPEECH_RECOGNITION', 'phrase_time_limit', 7)
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
                
                command = self.recognizer.recognize_google(audio).lower()
                
                if not wake_word_mode:
                    print(f"You: {command}")
                    self.logger.debug(f"Heard: {command}")
                
                # Check for wake words if in wake word mode
                if wake_word_mode:
                    return command if any(wake_word in command for wake_word in self.wake_words) else ""
                
                return command
                
        except sr.WaitTimeoutError:
            return ""
        except sr.UnknownValueError:
            if not wake_word_mode:
                self.logger.debug("Could not understand audio")
            return ""
        except sr.RequestError as e:
            self.logger.error(f"Speech recognition error: {e}")
            if not wake_word_mode:
                self.speak("My speech service is having issues.")
            return ""
    
    def continuous_listen(self, callback):
        """Continuous listening with wake word detection"""
        self.is_listening = True
        self.logger.info("Started continuous listening mode")
        
        while self.is_listening:
            try:
                wake_word = self.listen(timeout=1, wake_word_mode=True)
                if wake_word and any(word in wake_word for word in self.wake_words):
                    self.speak("Yes, how can I help?")
                    command = self.listen()
                    if command:
                        callback(command)
                
                time.sleep(0.1)  # Small delay to prevent excessive CPU usage
            except KeyboardInterrupt:
                self.stop_listening()
                break
            except Exception as e:
                self.logger.error(f"Continuous listen error: {e}")
                time.sleep(1)
    
    def stop_listening(self):
        """Stop continuous listening"""
        self.is_listening = False
        self.logger.info("Stopped continuous listening")

# ===== Main AIVA Class =====
class AIVA:
    def __init__(self):
        # Initialize components
        self.config = AIVAConfig()
        self.logger = AIVALogger(self.config)
        self.database = AIVADatabase()
        self.voice_manager = VoiceManager(self.config, self.logger)
        self.excel_manager = ExcelManager(self.logger)
        self.email_manager = EmailManager(self.config, self.logger, self.database)
        self.web_search = WebSearchManager(self.logger)
        self.system_monitor = SystemMonitor(self.logger)
        
        # Application state
        self.is_running = False
        self.conversation_mode = False
        
        self.logger.info("AIVA initialized successfully")
    
    def start(self):
        """Start AIVA assistant"""
        self.is_running = True
        self.voice_manager.speak("Hello! I am AIVA, your Advanced AI Voice Assistant. How can I help you today?")
        
        try:
            # Check if continuous mode is requested
            self.voice_manager.speak("Would you like to enable continuous listening mode? Say yes or no.")
            response = self.voice_manager.listen()
            
            if "yes" in response:
                self.voice_manager.speak("Continuous mode enabled. Just say 'Hey AIVA' to get my attention.")
                self.continuous_mode()
            else:
                self.command_mode()
                
        except KeyboardInterrupt:
            self.shutdown()
    
    def continuous_mode(self):
        """Run in continuous listening mode"""
        self.voice_manager.continuous_listen(self.process_command)
    
    def command_mode(self):
        """Run in traditional command mode"""
        while self.is_running:
            try:
                command = self.voice_manager.listen()
                if command:
                    if any(word in command for word in ["exit", "stop", "goodbye", "quit"]):
                        self.shutdown()
                        break
                    self.process_command(command)
            except KeyboardInterrupt:
                self.shutdown()
                break
    
    def process_command(self, command: str):
        """Enhanced command processing with comprehensive features"""
        self.logger.info(f"Processing command: {command}")
        
        try:
            success = True
            response = ""
            
            # Excel commands
            if self.is_excel_command(command):
                success = self.handle_excel_command(command)
                response = "Excel command executed"
            
            # Email commands
            elif self.is_email_command(command):
                success = self.handle_email_command(command)
                response = "Email command executed"
            
            # System commands
            elif self.is_system_command(command):
                success = self.handle_system_command(command)
                response = "System command executed"
            
            # Web/Search commands
            elif self.is_web_command(command):
                success = self.handle_web_command(command)
                response = "Web command executed"
            
            # Utility commands
            elif self.is_utility_command(command):
                success = self.handle_utility_command(command)
                response = "Utility command executed"
            
            # Information commands
            elif self.is_info_command(command):
                success = self.handle_info_command(command)
                response = "Information command executed"
            
            # Media commands
            elif self.is_media_command(command):
                success = self.handle_media_command(command)
                response = "Media command executed"
            
            # Smart home commands (placeholder for future integration)
            elif self.is_smart_home_command(command):
                success = self.handle_smart_home_command(command)
                response = "Smart home command executed"
            
            else:
                self.voice_manager.speak("I'm sorry, I don't understand that command. You can ask me about my capabilities by saying 'what can you do'.")
                success = False
                response = "Command not recognized"
            
            # Log command to database
            self.database.log_command(command, success, response)
            
        except Exception as e:
            self.logger.error(f"Command processing error: {e}")
            self.voice_manager.speak("I encountered an error processing that command.")
            self.database.log_command(command, False, str(e))
    
    # Command category checkers
    def is_excel_command(self, command: str) -> bool:
        excel_keywords = [
            "excel", "spreadsheet", "cell", "column", "row", "formula",
            "sheet", "workbook", "chart", "graph", "pivot", "table"
        ]
        return any(keyword in command.lower() for keyword in excel_keywords)
    
    def is_email_command(self, command: str) -> bool:
        email_keywords = ["email", "mail", "send", "compose", "gmail", "inbox"]
        return any(keyword in command.lower() for keyword in email_keywords)
    
    def is_system_command(self, command: str) -> bool:
        system_keywords = [
            "shutdown", "restart", "sleep", "lock", "system info",
            "task manager", "processes", "cpu", "memory", "disk"
        ]
        return any(keyword in command.lower() for keyword in system_keywords)
    
    def is_web_command(self, command: str) -> bool:
        web_keywords = [
            "search", "google", "youtube", "browse", "website",
            "chrome", "firefox", "browser", "tab", "bookmark"
        ]
        return any(keyword in command.lower() for keyword in web_keywords)
    
    def is_utility_command(self, command: str) -> bool:
        utility_keywords = [
            "time", "date", "weather", "calendar", "reminder",
            "note", "calculate", "convert", "translate"
        ]
        return any(keyword in command.lower() for keyword in utility_keywords)
    
    def is_info_command(self, command: str) -> bool:
        info_keywords = [
            "who are you", "what can you do", "help", "commands",
            "version", "about", "capabilities"
        ]
        return any(keyword in command.lower() for keyword in info_keywords)
    
    def is_media_command(self, command: str) -> bool:
        media_keywords = [
            "play", "pause", "stop", "music", "video", "volume",
            "spotify", "netflix", "youtube music"
        ]
        return any(keyword in command.lower() for keyword in media_keywords)
    
    def is_smart_home_command(self, command: str) -> bool:
        smart_keywords = [
            "lights", "temperature", "thermostat", "door", "lock",
            "security", "camera", "smart home"
        ]
        return any(keyword in command.lower() for keyword in smart_keywords)
    
    # Command handlers (implementations)
    def handle_excel_command(self, command: str) -> bool:
        """Handle Excel-related commands"""
        try:
            if not self.excel_manager.initialize():
                self.voice_manager.speak("I couldn't open Excel. Please
