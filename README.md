# EnergyReview - Alarm Dashboard

A premium React dashboard to analyze network energy alarms from Excel files.

## Features
- **Excel Analysis**: Automatically parses the first 5 subsheets of your alarms excel.
- **Site Mapping**: Detects site codes from Column F.
- **Weekly History**: Stores snapshots in Firebase to compare performance across weeks.
- **Premium UI**: Modern design with dark mode support, glassmorphism, and smooth animations.

## Setup Instructions

### 1. Firebase Setup (Free Tier)
1. Go to [Firebase Console](https://console.firebase.google.com/).
2. Create a new project named `Energy-Review`.
3. Add a "Web App" to your project.
4. Copy the `firebaseConfig` object.
5. Open `src/lib/firebase.js` and replace the placeholders with your credentials.
6. In Firebase Console, go to **Firestore Database** and click **Create Database**.
7. Set the rules to "Test Mode" (or configure secure rules) to allow data storage.

### 2. Local Development
1. Clone the repository.
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the development server:
   ```bash
   npm run dev
   ```

### 3. Usage
- Drag and drop your alarms Excel file into the dashboard.
- The system will process the first 5 subsheets.
- Data will be saved to Firebase and appear in the "Recent Uploads" section for comparison.

## Tech Stack
- **Frontend**: React, Tailwind CSS, Vite
- **Storage**: Firebase Firestore
- **Analysis**: XLSX.js
- **Charts**: Recharts
- **Icons**: Lucide React
- **Animations**: Framer Motion
