# AutoCapture
## Screenshot → PDF Appender

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

- Autoscroll mode  
  Automatically scrolls and captures content continuously with adjustable scroll amount, direction, and loop estimation.

- AutoStitch mode (seamless single-page output)  
  Combines multiple captures into a single continuous page by intelligently removing overlapping regions.

- Merge PDF to single page  
  Convert an existing multi-page PDF into a single vertically merged page.

- Session autosave  
  All settings are automatically saved and restored on next launch.

- Import / Export progress  
  Save your configuration and reload it later or on another system.  

- Convenience UX
  - Press Esc anytime (even outside the app) to safely stop automation.
  - Approximate loop confirmation - Prompts user after estimated loops in autoscroll to continue or stop safely.
  - Auto PDF conflict handling - Choose between append, overwrite, or auto-rename when file already exists.
  - Selectable capture region - Drag and define the exact area of the screen to capture.
  - Preview before saving - Accept or reject clipboard captures.
  - Optional image upscaling (2×) - Improves clarity before adding to PDF.
  - Always-on-top mode - Keeps the app accessible while working.

---

## Tips 
**Lazy to read? [Go to Quick Tips](#quick-tips)**

- Disable Windows animations for better capture reliability  
  (Settings → Accessibility → Animation Effects → Off)

- Use `Win + Shift + S` for the fastest workflow

- AutoStitch works best when consecutive captures have slight overlap.  
  Recommended: ~3 captures per page for clean stitching.

- For AutoScroll:
  - Lower scroll units = more overlap (safer stitching)
  - Higher scroll units = faster but risk gaps

- Do not interact with the screen during Autoscroll/Autonext runs.

- AutoStitch processes after pressing Stop — large captures may take time (especially 50+ frames).

- If stitching results look misaligned:
  - Reduce scroll units
  - Ensure consistent zoom level
  - Avoid dynamic content (animations, loading elements)

- Clipboard preview ensures you don’t accidentally append wrong captures.

- Global Esc is preferred over repeatedly clicking Stop.

- Upscaling improves readability but increases file size.

- Approx loops in Autoscroll is only an estimate — you can continue capturing beyond it.

- Merge tool loads entire PDF into memory — large PDFs may take time.
- The app automatically saves progress to a local `.autocapture` file  
- You can manually export/import this file for backup or sharing setups
- Press Esc to stop automation.  
  Avoid pressing repeatedly — wait a moment for response.  
  If it does not respond, press Esc again once.

- Hover over (?) icons in the app for detailed explanations of each setting.

- Some fields are disabled based on selected modes (Autoscroll / Autonext).  
  Enable the relevant option to unlock them.

- If results are not as expected, retry with the same settings after minor adjustments.

## Quick Tips
**Please do these before you start**
- Disable Windows animations before use

- Use Win + Shift + S for fastest capture

- Keep zoom level consistent for clean results

- For AutoScroll:
  Lower scroll = safer overlap  
  Higher scroll = faster but risky

- Do not touch mouse/keyboard during automation

- Press Esc once to stop (wait a moment before pressing again)

- AutoStitch needs overlap (~3 captures per page works best)

- Large AutoStitch runs may take time — be patient

- If output looks wrong:
  adjust scroll OR retry capture

- Settings auto-save; use export for backup

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
   - Supports precise click automation using recorded screen coordinates  
   - Ideal for fixed-layout navigation (e.g., next button, slideshow arrows)

   Autoscroll Mode
   - Define capture region and scroll amount (trial and error)   
   - Set approximate loop count (as to how many loops might be needed to finish the entire document)   
   - App captures and navigates automatically 

   - Optional: Enable AutoStitch to merge all captures into one seamless page  
   - App will prompt after approximate loops to continue or stop  
   - Final stitched page is generated when you press Stop  

   Automerge separate pages into 1 page to create 1 page pdf? (with file picker)

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

- The app automatically saves progress to a local `.autocapture` file  
- You can manually export/import this file for backup or sharing setups
---

## Build Executable

Using PyInstaller:  

`python -m PyInstaller AutoCapture.spec` 

The compiled executable will be available in the dist/ folder.

---