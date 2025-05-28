# AIVA - Advanced AI Voice Assistant

![AIVA Logo](https://img.shields.io/badge/AIVA-Voice%20Assistant-blue?style=for-the-badge&logo=robot)

AIVA (Artificial Intelligence Virtual Assistant) is a comprehensive voice-controlled assistant built in Python that can help you automate various tasks on your computer through natural voice commands.

## 🌟 Features

### 📧 Email Management
- **Smart Email Composition**: Send emails through Gmail with voice commands
- **Template System**: Create and use custom email templates
- **Intelligent Parsing**: Natural language processing for email recipients and subjects
- **Auto-send Capability**: Automatically send emails or draft them for review

### 📊 Excel Automation
- **Complete Spreadsheet Control**: Create, edit, and format Excel spreadsheets
- **Advanced Formulas**: Generate SUM, AVERAGE, MAX, MIN, and COUNT formulas
- **Chart Creation**: Create various types of charts from your data
- **Conditional Formatting**: Apply visual formatting based on cell values
- **Sheet Management**: Create, rename, and manage multiple worksheets

### 🌐 Web Browsing
- **Intelligent Search**: Search Google, YouTube, and other websites
- **Browser Control**: Open tabs, close tabs, scroll, and refresh pages
- **Quick Access**: Launch popular websites with simple voice commands
- **Bookmark Management**: Save and access your favorite sites

### 🖥️ System Control
- **Power Management**: Shutdown, restart, sleep, and lock your computer
- **System Monitoring**: Check CPU usage, memory, and disk space
- **Process Management**: View and manage running applications
- **Application Launcher**: Open any installed application

### 🎵 Media Control
- **Music Playback**: Control Spotify, YouTube Music, and other players
- **Volume Control**: Adjust system volume with voice commands
- **Media Navigation**: Play, pause, skip, and control media playback

### 🏠 Smart Home Integration (Coming Soon)
- **IoT Device Control**: Control smart lights, thermostats, and security systems
- **Scene Management**: Create and activate home automation scenes
- **Voice-Activated Security**: Arm/disarm security systems with voice

### 🛠️ Utility Features
- **Time & Date**: Get current time, date, and set reminders
- **Weather Information**: Real-time weather updates for any location
- **Unit Conversion**: Convert between different units of measurement
- **Calculator**: Perform mathematical calculations
- **Note Taking**: Create and manage voice notes

## 🚀 Quick Start

### Prerequisites

- **Python 3.8+** (Python 3.9+ recommended)
- **Windows OS** (for Excel integration via COM)
- **Microphone** (for voice input)
- **Speakers/Headphones** (for voice output)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/aiva-voice-assistant.git
   cd aiva-voice-assistant
   ```

2. **Create a virtual environment**
   ```bash
   python -m venv aiva_env
   aiva_env\Scripts\activate  # Windows
   # or
   source aiva_env/bin/activate  # macOS/Linux
   ```

3. **Install required packages**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run AIVA**
   ```bash
   python aiva_main.py
   ```

## 📋 Requirements

Create a `requirements.txt` file with the following dependencies:

```txt
speechrecognition==3.10.0
pyttsx3==2.90
pyautogui==0.9.54
keyboard==0.13.5
pyjokes==0.6.0
pywin32==306
requests==2.31.0
psutil==5.9.5
configparser==6.0.0
```

## 📁 Project Structure

```
aiva-voice-assistant/
│
├── aiva_main.py              # Main application entry point
├── config/
│   └── aiva_config.ini       # Configuration settings
├── data/
│   └── aiva.db              # SQLite database for logs and preferences
├── logs/
│   └── aiva.log             # Application logs
├── templates/
│   └── email_templates.json  # Email templates
├── docs/
│   ├── API_REFERENCE.md     # API documentation
│   ├── COMMANDS.md          # Complete command list
│   └── TROUBLESHOOTING.md   # Common issues and solutions
├── tests/
│   ├── test_voice.py        # Voice management tests
│   ├── test_email.py        # Email functionality tests
│   └── test_excel.py        # Excel automation tests
├── requirements.txt         # Python dependencies
├── setup.py                # Package setup script
├── LICENSE                 # MIT License
└── README.md              # This file
```

## 🎮 Usage Examples

### Basic Commands

```
"Hey AIVA, what time is it?"
"Open Gmail"
"Search for Python tutorials"
"What's the weather like?"
"Tell me a joke"
```

### Email Commands

```
"Send an email to john@example.com regarding meeting tomorrow"
"Compose email to team@company.com about project update"
"Send mail using template meeting to boss@company.com"
```

### Excel Commands

```
"Open Excel"
"Write Sales Data in cell A1"
"Fill column B with January, February, March"
"Sum formula in cell C10 for C1 to C9"
"Create a chart from A1 to C10"
```

### System Commands

```
"Show system information"
"What processes are running?"
"Lock my computer"
"Shutdown in 5 minutes"
```

## ⚙️ Configuration

AIVA uses a configuration file located at `config/aiva_config.ini`. You can customize:

### Voice Settings
```ini
[VOICE]
rate = 150              # Speech rate (words per minute)
volume = 0.9           # Voice volume (0.0 to 1.0)
voice_index = 0        # Voice selection index
```

### Email Settings
```ini
[EMAIL]
default_domain = gmail.com
auto_send = true
send_delay = 5
```

### Application Settings
```ini
[APPLICATIONS]
chrome_path = C:/Program Files/Google/Chrome/Application/chrome.exe
excel_auto_open = false
browser_delay = 3
```

## 🔧 Advanced Features

### Custom Email Templates

Create custom email templates in the database:

```python
# Example: Adding a new template
aiva.database.save_email_template(
    name="meeting_request",
    subject="Meeting Request - {topic}",
    body="Dear {recipient},\n\nI would like to schedule a meeting regarding {topic}..."
)
```

### Wake Word Activation

Enable hands-free operation with wake words:
- "Hey AIVA"
- "OK AIVA"
- "AIVA"

### Continuous Listening Mode

AIVA can run in the background and respond to wake words without manual activation.

## 🧪 Testing

Run the test suite to ensure everything works correctly:

```bash
# Run all tests
python -m pytest tests/

# Run specific test categories
python -m pytest tests/test_voice.py
python -m pytest tests/test_email.py
python -m pytest tests/test_excel.py
```

## 🐛 Troubleshooting

### Common Issues

1. **Microphone not working**
   - Check microphone permissions
   - Ensure microphone is set as default input device
   - Test with Windows Voice Recorder

2. **Excel commands not working**
   - Ensure Microsoft Excel is installed
   - Run Python as administrator for COM object access
   - Check if Excel is already running

3. **Speech recognition errors**
   - Check internet connection (Google Speech API requires internet)
   - Speak clearly and at moderate pace
   - Reduce background noise

4. **Email not sending**
   - Ensure you're logged into Gmail in your default browser
   - Check popup blocker settings
   - Verify internet connection

### Performance Optimization

- **Reduce CPU usage**: Increase listening timeout intervals
- **Improve accuracy**: Train with your voice patterns
- **Faster startup**: Enable application pre-loading

## 🔒 Privacy & Security

- **Local Processing**: Most operations happen locally on your machine
- **No Data Storage**: Voice commands are not permanently stored
- **Secure Communications**: HTTPS for all web requests
- **Permission-Based**: Only accesses applications you explicitly allow

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### Development Setup

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes and add tests
4. Ensure all tests pass: `python -m pytest`
5. Submit a pull request

### Code Style

We use:
- **Black** for code formatting
- **Flake8** for linting
- **Type hints** for better code documentation
- **Docstrings** for all public methods

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🚧 Roadmap

### Version 2.0 (Coming Soon)
- [ ] Multi-language support
- [ ] Custom wake word training
- [ ] Mobile app companion
- [ ] Cloud synchronization
- [ ] Plugin system for third-party integrations

### Version 2.5 (Future)
- [ ] Machine learning for personalized responses
- [ ] Advanced natural language understanding
- [ ] Integration with popular productivity tools
- [ ] Cross-platform support (macOS, Linux)

## 📞 Support

- **Documentation**: Check our [Wiki](https://github.com/yourusername/aiva-voice-assistant/wiki)
- **Issues**: Report bugs on [GitHub Issues](https://github.com/yourusername/aiva-voice-assistant/issues)
- **Discussions**: Join our [GitHub Discussions](https://github.com/yourusername/aiva-voice-assistant/discussions)
- **Email**: support@aiva-assistant.com

## 🌟 Acknowledgments

- **SpeechRecognition** library for voice input processing
- **pyttsx3** for text-to-speech functionality
- **Microsoft Office** for Excel automation capabilities
- **Google Speech API** for accurate speech recognition
- **PyAutoGUI** for system automation

## 📊 Stats

![GitHub stars](https://img.shields.io/github/stars/yourusername/aiva-voice-assistant?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/aiva-voice-assistant?style=social)
![GitHub issues](https://img.shields.io/github/issues/yourusername/aiva-voice-assistant)
![GitHub license](https://img.shields.io/github/license/yourusername/aiva-voice-assistant)

---

**Made with ❤️ by the AIVA Team**

*AIVA - Making your computer truly intelligent, one voice command at a time.*
