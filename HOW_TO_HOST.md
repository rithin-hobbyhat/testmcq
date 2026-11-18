# 📱 MCQ Practice Arena - Online Hosting & Smartphone Guide

This application has been converted into a high-performance, mobile-responsive **MCQ Practice Arena** that can be hosted online for free and played on smartphones, tablets, and computers.

---

## ⚡ Quick Start (Instant Play on your Smartphone Right Now)

If your computer and smartphone are on the **same Wi-Fi or Mobile Hotspot**:

1. Double-click **`Launch_Practice_Web.bat`** on your computer.
2. The terminal will display your computer's local IP address, for example:
   ```
   💻 On this PC:
      Open in Browser: http://localhost:8000

   📱 On your Smartphone:
      Open in Chrome/Safari: http://192.168.1.X:8000
   ```
3. Type the **Smartphone URL** into Chrome or Safari on your phone.
4. You are ready to practice all 5,700+ MCQs immediately!

---

## 🌐 Free Online Hosting (Play From Anywhere in the World)

Because the web application is built with modern, ultra-fast client-side code, it requires **no database configuration or backend servers** to host. The entire `web` folder is 100% production-ready.

### 🚀 Method 1: Netlify Drop (Easiest - 30 Seconds, No Coding)
1. Go to **[https://app.netlify.com/drop](https://app.netlify.com/drop)** (Sign up for a free Netlify account).
2. Simply **drag and drop** the **`web`** folder from this project onto the Netlify webpage.
3. Netlify will instantly give you a free public link (e.g. `https://my-mcq-practice.netlify.app`).
4. Open that link on your smartphone from anywhere with mobile data (4G/5G)!

---

### 🚀 Method 2: GitHub Pages (100% Free Forever)
1. Initialize git and push this project to a new repository on your GitHub account.
2. In your GitHub repository, click **Settings** > **Pages**.
3. Under **Branch**, select `main` (or `master`) and set folder to `/web` (or root if you put web files in the repository root).
4. Click **Save**. Within 1 minute, GitHub will give you a live URL:
   `https://<your-username>.github.io/<repo-name>/`

---

### 🚀 Method 3: Vercel (1-Click Deployment)
1. Go to **[https://vercel.com](https://vercel.com)**.
2. Import your GitHub repository or use Vercel CLI.
3. Set the Root Directory to `web`.
4. Click **Deploy**. Vercel will give you a lightning-fast global CDN link!

---

## 📲 Install as an App on Smartphone (PWA)

The app is built as a **Progressive Web App (PWA)**:
- **On Android (Chrome)**: Tap the 3 dots menu `⋮` > tap **"Add to Home screen"** or **"Install App"**.
- **On iPhone (Safari)**: Tap the Share button `⎋` > tap **"Add to Home Screen"**.

This installs the MCQ Practice Arena onto your home screen with a dedicated app icon and full-screen view (no browser address bar), and it can even work **100% offline** once cached!

---

## 🎯 Features Included in the Practice App

1. **Multi-Subject Selection**:
   - Practice questions from a single subject or select multiple subjects simultaneously (e.g., *AMVI* + *Transmission System* + *Chassis*).
   - "Select All" and "Clear All" shortcuts.
   - Live question count counter.

2. **Order Modes**:
   - 🔀 **Random Order**: Shuffles all selected questions for a realistic mock exam experience.
   - 🔢 **As Per Order**: Asks questions sequentially in the exact order of the original Excel database.

3. **Immediate Answer Feedback & Description**:
   - Selecting a wrong answer immediately turns the choice red (❌), highlights the correct answer in green (✔), and smoothly displays the **full explanation / description** from your database!

4. **Live Score Tracking**:
   - Real-time score display: `Score: [Right Answers] / [Total Attended]` (e.g. `12 / 15`).
   - Live accuracy percentage (e.g. `80.0%`).
   - Visual counters for `Right`, `Wrong`, and `Remaining` questions.
   - Animated progress bar.

5. **End-of-Session Performance Report & Review**:
   - Comprehensive score card with percentage and grading.
   - Question review with filters: "Show All", "Mistakes Only", and "Correct Only".
   - **Retest Mistakes Only** button to immediately drill down on questions you missed!

6. **Updating Questions**:
   - Whenever you add new questions in the Excel sheet or through the Question Bank Studio (`app.py`), simply run **`py export_data.py`** or click **"🎯 Practice MCQs (Web/Phone)"** inside the desktop studio, and the web app updates automatically.
