📌 Project Overview
Workplace Monitoring with Automated Alerts is a real-time employee detection system using deep learning and computer vision. The system uses the YOLOv3 algorithm to detect people from webcam footage and triggers alerts if no one is detected for a specified time.

🚀 Key Features
Real-time person detection using YOLOv3 and OpenCV
Automatic alerts via email with attached image
Excel logging of detection events with timestamps
Audible buzzer and text-to-speech notifications
Optimized for laptops with built-in webcams
Simple, scalable, and privacy-considerate design

🎯 Objectives
Detect employee presence in real time
Trigger alerts after 60 seconds of absence
Log events for auditing
Provide an affordable and scalable solution
Ensure privacy-friendly monitoring

🧠 Technologies Used
YOLOv3 – Object detection
OpenCV – Webcam integration and image processing
Python – Core logic and automation
smtplib – Email alert system
openpyxl – Excel logging
pyttsx3 – Text-to-speech alert
winsound – Audible alert

🏗️ System Architecture
Webcam captures real-time video
YOLOv3 detects presence of a person
If no person is detected:
Image is saved
Email is sent with snapshot
Excel sheet logs absence
Buzzer and voice alert triggered
Resets detection when employee is back

📈 Output Snapshots
✅ Employee Detected – Green box drawn on screen
❌ Employee Not Detected – Alert triggered
📧 Email Notification – With image of absence
📊 Excel Logs – Date, Time, Status stored

⚙️ Requirements
Software
Windows 10+
Python 3.7.9
Libraries: opencv-python, numpy, smtplib, pyttsx3, openpyxl
Hardware
Minimum: 8GB RAM, Intel i5 CPU
Webcam (built-in or external)
(Optional) GPU for faster performance

📚 Acknowledgments
Developed by:
Sahana Patil H V – 01JST21IS042
Chaitra N – 01JST22UIS400
Reethu M – 01JST22UIS407
Venkatesh Gudikoti – 01JST22UIS410

Under the guidance of Prof. Lavanya M S,
Department of Information Science and Engineering,
JSS Science and Technology University, Mysuru.
