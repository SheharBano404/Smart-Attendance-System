# Smart-Attendance-System
A Python-based face recognition and voice-controlled attendance system built with OpenCV, Tkinter, and speech APIs. Features auto-registration, admin vault, and CSV/JSON logging.

# Abstract
The Smart Attendance System automates attendance using face recognition, voice commands, and a Tkinter interface. It captures faces via webcam, logs attendance to CSV, and provides an admin vault for viewing records, blocking users, and managing registrations.

# Features
Face detection (Haar Cascade) and recognition (LBPH)

Auto-registration for unknown users
Admin vault with attendance view, block/unblock, and admin roll-call
Voice control with wake word “hello”
Duplicate prevention per day
Playful themed feedback

# Core Stack
OpenCV (Haar Cascade + LBPH)
Tkinter
PIL (ImageTk)
SAPI (Windows Speech)
speech_recognition (Google)
CSV/JSON

# Setup Instructions
Clone repo
Install dependencies (pip install opencv-contrib-python numpy pillow pywin32 SpeechRecognition)
Run python smart_attendance.py in VS Code

# First-Time Steps
Admin login (username: sheharbano404/passkey: snape)
Capture admin face (press S)
Train model (auto after capture/registration)
Start attendance (GUI button or voice command)
You don't need to worry, the system will guide you the steps to explore different functionalities. 

# Data Files
faces/<Name>/*.jpg
faces/Admin/admin.jpg
trainer.yml
label_map.json, id_map.json
attendance.csv
blocked.json
