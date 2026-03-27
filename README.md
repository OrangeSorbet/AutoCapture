# Screenshot → PDF Appender

A lightweight desktop tool to capture screenshots and append them directly into a single PDF. Useful for notes, documentation, and tutorials.

---

## Features

- Screenshot to PDF (instant append)  
  Capture screenshots and automatically add them as new pages in a PDF.

- Clipboard monitoring  
  Detects screenshots copied via `Win + Shift + S` and allows preview before saving.

- Custom hotkey capture  
  Define a hotkey to capture a specific screen region instantly.

- Autonext mode (automation)  
  Automates:
  - Capture → Click → Capture loop  
  Useful for scrolling content, slides, or paginated documents.

- Selectable capture region  
  Drag and define the exact area of the screen to capture.

- Preview before saving  
  Accept or reject clipboard captures.

- Optional image upscaling (2×)  
  Improves clarity before adding to PDF.

- Always-on-top mode  
  Keeps the app accessible while working.

---

## How It Works

1. Set:
   - PDF name  
   - Save location  

2. Choose a mode:

   Manual Mode
   - Use `Win + Shift + S`
   - Preview appears → Accept → Appends to PDF

   Hotkey Mode
   - Set hotkey and capture area  
   - Press hotkey → Screenshot is saved instantly

   Autonext Mode
   - Define capture region and click position  
   - Set loop count  
   - App captures and navigates automatically  

---

## Getting Started (Developers)

### 1. Clone the Repository

`git clone https://github.com/OrangeSorbet/AutoCapture.git`
`cd AutoCapture`

### 2. Create Virtual Environment

`python -m venv venv`

### 3. Activate It

Windows:
`venv\Scripts\activate`

### 4. Install Dependencies

`pip install -r requirements.txt`

---

## Build Executable

Using PyInstaller:

`python -m PyInstaller AutoCapture.spec`

The compiled executable will be available in the dist/ folder.

---

## Tips

- Disable Windows animations for better capture reliability  
  (Settings → Accessibility → Animation Effects → Off)

- Use `Win + Shift + S` for the fastest workflow