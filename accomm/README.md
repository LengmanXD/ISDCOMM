# ACComm – Access Control Commissioning Tool

A mobile-first web app for on-site engineers to build commissioning documentation on the go.

## Features

- **Device Inventory** – Log all access control hardware (ACUs, readers, locks, intercoms, cameras, barriers)
- **Test Records** – PASS/FAIL/N/A per test type with auto-calculated overall result
- **Cable Schedule** – Full cable log with from/to, type, cores, length, route
- **Photo Log** – On-site photo capture with reference tags
- **Auto-save** – All data saved automatically to browser localStorage
- **Multi-project** – Switch between multiple site projects
- **Excel Export** – One-tap export to .xlsx with 6 sheets

---

## Setup

### 1. Install Node.js
Download from https://nodejs.org (LTS version recommended)

### 2. Install dependencies
Open a terminal in this folder and run:
```bash
npm install
```

### 3. Start the development server
```bash
npm run dev
```

Then open http://localhost:5173 in your browser (or on your phone if on the same network).

---

## Build for production (optional)

To create a production build you can host on any web server:
```bash
npm run build
```
Output goes into the `dist/` folder. You can host this on any static hosting service (Netlify, Vercel, GitHub Pages, your own server, etc.).

---

## Using on mobile

While running `npm run dev`, find your computer's local IP address (e.g. 192.168.1.10) and open:
```
http://192.168.1.10:5173
```
on your phone's browser. Both devices must be on the same Wi-Fi network.

For permanent mobile access, deploy the production build to a hosting service.

---

## Data storage

All project data is stored in your **browser's localStorage** — it stays on the device and persists between sessions. Nothing is sent to any server.

> Note: Clearing browser site data will erase saved projects. Export your XLSX regularly as a backup.

---

## Project structure

```
accomm/
├── index.html
├── package.json
├── vite.config.js
└── src/
    ├── main.jsx      # React entry point
    └── App.jsx       # Full application
```
