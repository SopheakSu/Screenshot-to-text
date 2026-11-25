SCREEN TO EXCEL  -  OCR tool for the alarm report
=================================================

GET THE .EXE  -  OPTION A: on your PC (one time, about 3-5 minutes)
 1. Install Python 3.10 or newer from https://www.python.org/downloads/
    (on the first install screen tick "Add python.exe to PATH").
 2. Unzip this folder, then double-click  build_exe.bat
 3. When it finishes, your app is in the "dist" folder:  ScreenToExcel.exe
    Copy that one file anywhere (Desktop) - it is all you need, and other PCs
    don't need Python installed.

GET THE .EXE  -  OPTION B: no Python needed (GitHub builds it for you)
 1. Make a free account at github.com and create a new repository.
 2. "Add file" > "Upload files", drag in everything in this folder
    (including the hidden .github folder), then "Commit changes".
 3. Open the "Actions" tab, wait for the green tick (about 3 minutes),
    click the run, and download "ScreenToExcel-exe" at the bottom.
    Unzip it to get ScreenToExcel.exe.

DAILY USE
 1. Open ScreenToExcel.exe (leave it running).
 2. Take a screenshot of the rows with any app you like (Snipping Tool,
    Win+Shift+S, Lightshot...) so the image is copied, then click the app
    and press Ctrl+V.  Or click "Open image..." to load a saved picture.
 3. The table shows exactly the rows in that screenshot (1 row in the
    screenshot = 1 row in the result) and they are copied automatically.
 4. In Excel click the first empty cell under "Code" and press Ctrl+V.

OUTPUT COLUMNS (default, same order as your Excel sheet)
 Code | Name | Contact (blank) | Area | Zone | Type
   01-000486 | SAKURA Recycle Shop Japan 271 | | 2 | GF-Ent-DP1-3 | Intruder
 - Area comes from T01-0X-1  (T01-02-1 -> 2).
 - "-____" after the code is removed automatically.
 - Tick/untick columns in Options (e.g. add Time). Settings are remembered.

TIPS
 - Double-click any cell to fix an OCR mistake; Delete removes rows.
 - Tick "Keep old rows" if you want to snip several screens and paste
   them all at once (off by default: each new screenshot replaces the table).
 - "Skip duplicates" ignores rows you already captured.
 - OCR engine "auto" uses Windows' built-in OCR (needs the English language,
   which most PCs have: Settings > Time & Language > Language).
   Optional fallback: install Tesseract OCR (UB-Mannheim build) and choose
   "tesseract".
 - Always glance over the codes before saving - OCR can confuse 0/O, 1/I, 3/8.
