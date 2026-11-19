# ⚡ MCQ Practice Arena & Question Bank Studio (Version 3.0)

A high-performance, mobile-responsive **Online MCQ Practice Arena** and **Question Bank Authoring Studio** featuring the innovative **Pocket Study Mode**, pure high-contrast **Black & White Dark Mode**, multi-subject selection, and real-time performance analytics.

---

## 🌐 Live Online Practice

Play and practice over 5,700+ questions directly in your browser on smartphones, tablets, or computers:

* **Primary Live Arena:** [https://rithin-hobbyhat.github.io/hobbyhat/mcq/](https://rithin-hobbyhat.github.io/hobbyhat/mcq/)
* **Repository Mirror:** [https://rithin-hobbyhat.github.io/testmcq/](https://rithin-hobbyhat.github.io/testmcq/)

Installable as a **Progressive Web App (PWA)** for 1-tap fullscreen offline study on Android Chrome and iOS Safari.

---

## 🚀 Key Upgrades in Version 3.0

### 1. ⏱️ Innovative "Pocket Study Mode" (5-Minute Background Interval)
* **Hands-Free Study**: Keep Chrome minimized or your phone in your pocket while doing other daily activities.
* **Smart Alert Every 5 Minutes**:
  * **Device Vibration**: Distinct tactical dual-pulse vibration pattern (`[400ms, 200ms, 400ms, 200ms, 600ms]`).
  * **Audio Bell Chime**: Crystal-clear alert chime generated via the Web Audio API.
  * **Push Notification**: Displays the upcoming question topic directly on your phone's lock screen.
* **Auto-Standby on Correct Answer**: If you select the right answer, the question disappears after 1.2 seconds, your score is logged, and the app enters standby until the next 5-minute alert.
* **Deep Review on Mistake**: If you select the wrong answer, the detailed answer explanation and description are displayed immediately. Resume standby whenever you are ready.
* **Background Worker**: Powered by a dedicated Web Worker blob and visibility watchdog that prevents mobile OS tab sleeping and timer throttling.

### 2. 🖤 High-Contrast Pure Black & White Dark Mode
* Clean, distraction-free aesthetic with true `#000000` AMOLED blacks, crisp `#FFFFFF` typography, and `#18181B` zinc cards.
* Zero emojis — replaced with crisp, professional SVG vector iconography.
* Battery-saving and optimal for late-night or outdoor study.

### 3. 🎯 Multi-Subject Selection & Order Control
* Practice one subject or select multiple subjects simultaneously (e.g., AMVI, Transmission, Chassis, Electric & Hybrid Vehicles).
* **Sequential Order**: Follows the curated order of the question database.
* **Random Mode**: Shuffles questions across all selected subjects for realistic mock exam drills.
* Always runs all available questions for comprehensive preparation.

### 4. 🔒 Strict Interactive Rules
* **Locked Next Button**: The "Next" button remains disabled until an answer option has been selected.
* **Double-Click/Tap Explanation**: Double-clicking or double-tapping the correct answer option toggles the full explanation and description.
* **Focused Review**: Performance review screen filters directly into **Mistakes Only** and **Correct Only** to maximize revision efficiency.

---

## 💻 Desktop Question Bank Authoring Studio

Included in this repository is the full Windows Desktop GUI for authoring, editing, and managing MCQs directly into structured Excel sheets (`data/questions_bank.xlsx`).

### 🖱️ How to Run Locally:
1. **Desktop App**: Double-click `Launch_App.bat` or `Launch_Silent.vbs`.
2. **Local Web Server**: Double-click `Launch_Practice_Web.bat` to host on your local Wi-Fi and play on your smartphone at `http://<your-ip>:8000`.
3. **Data Sync**: Run `py export_data.py` to regenerate JSON web datasets from `questions_bank.xlsx`.

---

## 📁 Repository Structure

```
├── index.html              # Root redirect & fast loader for GitHub Pages
├── docs/                   # Production web build (GitHub Pages /docs target)
├── web/                    # MCQ Web Practice Arena source
│   ├── index.html          # Main web application
│   ├── styles.css          # Black & white dark theme stylesheet
│   ├── app.js              # Application logic, audio engine & pocket timer
│   ├── sw.js               # Service Worker (PWA caching & background notification)
│   ├── manifest.json       # PWA manifest
│   └── data/               # Subject-wise & master question datasets (JSON)
├── data/                   # Excel database (questions_bank.xlsx) & backups
├── app.py                  # Windows desktop authoring studio GUI
├── config.py               # Application configuration & symbol palette
├── export_data.py          # Excel to web JSON data exporter
├── models.py               # Data models and question schema
├── server.py               # Local multi-device HTTP server
├── Launch_App.bat          # Desktop app launcher
├── Launch_Practice_Web.bat # Local web practice launcher
└── README.md
```

---

## 🛠️ GitHub Pages Hosting for `rithin-hobbyhat/testmcq`

To host directly from this repository:
1. Go to **Settings** > **Pages** in this repository.
2. Under **Build and deployment** > **Branch**:
   - Choose `main` branch.
   - Select `/ (root)` or `/docs`.
3. Click **Save**.
4. The app is live at `https://rithin-hobbyhat.github.io/testmcq/`!
