# Energy Review - Master Validator

A high-performance, locally-hosted Electron + React desktop application designed to streamline, automate, and strictly audit Battery Discharge Test (BDT) reporting and Network Alarm (AUTIN) monitoring. 

## ⚡ Core Features

### 1. Alarm Analysis Engine (AUTIN)
- **Intelligent Excel Parsing**: Automatically ingests raw AUTIN Excel logs.
- **Dense Data Visualization**: Professional, high-density single-line table mode maximizing screen real estate.
- **Impacted Technology Tracking**: Visually highlights impacted layers (2G, 3G, 4G, 5G) during critical downtimes.
- **Image Export**: Built-in "Copy Table as Image" to instantly capture cleanly formatted tables for escalations and reporting.

### 2. Comprehensive BDT Auditing
- **Automated Validation**: Parses and strictly audits BDT excel forms against multiple engineering parameters:
  - Checks starting Amperes.
  - Validates sequence drops (Voltage should decrease/stay flat, Ampere logic).
  - Matches Rectifier logic with Battery discharge curves.
  - Calculates theoretical vs tested backup durations.
- **Asset Loss Detection**: Upload an **Old Summary** sheet to automatically cross-check `# of Modules`. If the current site shows fewer modules than historically recorded, the sheet is automatically flagged for "Asset Loss".
- **Current Summary Verification**: Cross-correlates the current BDT forms against a master "Current Summary" sheet to ensure data integrity and track reviewed vs unreviewed records.
- **Alarm Correlation**: Checks the exact time window of the BDT test against the loaded Alarm database to confirm power outage alarms actually triggered on-site.

### 3. Alarm Memory & Persistent Storage
- **Offline First**: All heavy data processing happens locally. No cloud databases required.
- **Alarm Database Memory**: Processed alarm datasets are securely saved locally into IndexedDB.
- **Re-Analysis**: Instantly re-run a BDT correlation against a specifically selected historical alarm database without needing to locate and re-upload the original raw files.

### 4. Dynamic Theme Console
- **Display Modes**: Instantly toggle between Light Mode and Night Mode.
- **Premium Color Themes**: Deeply customized CSS variable palettes ensuring excellent contrast ratios. Choose between:
  - **Default (Midnight Slate)**
  - **Sunset Orange**
  - **Emerald Green**
- **Persistent Settings**: Your theme preferences and custom BDT rule thresholds are saved in the app's Console settings and restore automatically on launch.

---

## 🚀 Setup & Execution

Since the project has been upgraded to an Electron-backed desktop app, Firebase and complex backend setups are no longer required.

### Local Development
1. Clone the repository and navigate to the project root.
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the Vite development server alongside the Electron app:
   ```bash
   npm run dev
   ```

### Building for Production
To package the app into a standalone Windows executable:
```bash
npm run build
```

---

## 🛠 Tech Stack
- **Frontend Framework**: React + Vite
- **Desktop Runtime**: Electron
- **Styling**: Tailwind CSS (Custom HSL Variable System)
- **Excel Parsing/Generation**: SheetJS (XLSX)
- **Local Persistence**: IndexedDB (via `localforage`) + `localStorage`
- **Icons**: Lucide React
- **Export utilities**: html2canvas
