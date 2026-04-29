# 🎓 Student Coding Profile Tracker

Automated Python script that fetches and updates student coding stats from LeetCode, HackerRank, and GitHub into an Excel sheet daily using Windows Task Scheduler.

---

## 📋 What It Does

- Reads student profile URLs from an Excel file
- Automatically fetches coding stats from **LeetCode**, **HackerRank**, and **GitHub**
- Updates the Excel sheet in-place with the latest data
- Runs daily on a schedule — no manual work needed

### Data it collects:

| Platform | Data Collected |
|---|---|
| **LeetCode** | Total problems solved, Easy, Medium, Hard count |
| **HackerRank** | Star ratings for Problem Solving, C, Python, Java, SQL |
| **GitHub** | Number of public repositories |

---

## 🖥️ Requirements

- Windows PC
- Internet connection
- Microsoft Excel or any software that can open `.xlsx` files
- Python installed on your PC

---

## ⚙️ Installation — Step by Step

### Step 1 — Install Python

1. Go to [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Click **Download Python** (latest stable version)
3. Run the installer
4. ✅ **Important:** Check **"Add Python to PATH"** before clicking Install
5. Click **Install Now**

To verify Python is installed, open Command Prompt and type:
```
python --version
```
You should see something like `Python 3.13.5`

---

### Step 2 — Download this Project

**Option A — Using Git:**
```bash
git clone https://github.com/yourusername/student-coding-profile-tracker.git
cd student-coding-profile-tracker
```

**Option B — Manual Download:**
1. Click the green **Code** button on this page
2. Click **Download ZIP**
3. Extract the ZIP file to a folder, e.g. `D:\CLG\excel automation`

---

### Step 3 — Set Up Virtual Environment

Open Command Prompt or PowerShell inside the project folder and run:

```bash
# Create virtual environment
python -m venv .venv

# Activate it
.venv\Scripts\activate
```

You should see `(.venv)` at the start of your terminal line — this means it is active.

---

### Step 4 — Install Required Packages

With the virtual environment active, run:

```bash
pip install openpyxl requests beautifulsoup4
```

---

### Step 5 — Set Up GitHub Token (Free)

The script needs a free GitHub token to fetch repo counts for 200+ students without hitting rate limits.

1. Go to [https://github.com/settings/tokens](https://github.com/settings/tokens)
2. Click **Generate new token (classic)**
3. Give it any name (e.g. `profile-tracker`)
4. Under scopes, select only ✅ **`public_repo`** — nothing else
5. Click **Generate Token**
6. **Copy the token immediately** — GitHub shows it only once

Now create your config file:

1. In the project folder, find `config.example.py`
2. Make a copy of it and rename the copy to `config.py`
3. Open `config.py` and replace `paste_your_token_here` with your copied token:

```python
GITHUB_TOKEN = "ghp_yourActualTokenHere"
```

> ⚠️ **Never share `config.py` with anyone or push it to GitHub — it contains your private token.**

---

### Step 6 — Prepare Your Excel File

Your Excel file must have these exact column headers in **Row 1**:

| Column | Header Name |
|---|---|
| A | Sno |
| B | Section |
| C | Roll Number |
| D | Name of the Student |
| E | HackerRank Profile URL |
| F | Problem Solving |
| G | C |
| H | Python |
| I | Java |
| J | SQL |
| K | LeetCode Profile URL |
| L | Total Leetcode Problems Solved |
| M | No.of Easy Problems Solved |
| N | No.of Medium Problems Solved |
| O | No.of Hard Problems Solved |
| P | GitHub Profile URL |
| Q | No.of Repos in GitHub |

Student data should start from **Row 2**.

URL formats expected:
- HackerRank: `https://www.hackerrank.com/profile/username`
- LeetCode: `https://leetcode.com/u/username/`
- GitHub: `https://github.com/username`

Place the Excel file inside the project folder.

---

## 🚀 Running the Script
> 📝 Note:
Before running the code, make sure the required Excel file is present in the same folder as the script. The code will not work if the Excel file is missing or placed in a different directory. Double-check that the Excel file is in the correct location before execution.
### Manual Run

1. Open terminal in the project folder
2. Activate the virtual environment:
```bash
.venv\Scripts\activate
```
3. Run the script:
```bash
python update_profiles.py
```
4. When prompted, type your Excel filename:
```
Enter the Excel filename (e.g. students.xlsx): ai_student_details.xlsx
```
5. Watch the progress in the terminal — it processes each student one by one
6. When done it saves the file automatically

> ⚠️ **Make sure the Excel file is closed before running the script.** Excel locks the file when open and the script cannot save to it.

---

### Testing on a Few Rows First (Recommended)

Before running on all 200+ students, test on 5 rows first:

1. Open `update_profiles.py`
2. Find this line:
```python
for row in range(2, total_rows + 1):
```
3. Change it to:
```python
for row in range(2, 7):
```
4. Run the script and verify the first 5 students look correct in Excel
5. Change the line back to `total_rows + 1` for the full run

---

## ⏰ Setting Up Daily Automatic Schedule (Windows Task Scheduler)

This makes the script run automatically every day at your chosen time.

1. Press **Windows + S** and search for **Task Scheduler** — open it
2. Click **Create Basic Task** on the right side
3. Give it a name: `Student Profile Tracker`
4. Click **Next** → Select **Daily** → Click **Next**
5. Set your preferred time (e.g. 8:00 AM) → Click **Next**
6. Select **Start a Program** → Click **Next**
7. In the **Program/Script** field, enter:
```
D:\CLG\excel automation\.venv\Scripts\python.exe
```
8. In the **Add arguments** field, enter:
```
update_profiles.py
```
9. In the **Start in** field, enter:
```
D:\CLG\excel automation
```
10. Click **Finish**

The script will now run automatically every day at your chosen time.

---

## ❗ Error Reference

If something goes wrong, the script writes these messages in the cell instead of crashing:

| Message | Meaning |
|---|---|
| `Profile Not Found` | Student's URL is wrong or account deleted |
| `Timeout` | Network was too slow, try again |
| `No Data` | Profile exists but returned no stats |
| `Error` | Something unexpected happened |

A detailed `error_log.txt` file is also created in the project folder after each run listing all failures with row number and student name.

---

## 📁 Project Structure

```
excel automation/
│
├── update_profiles.py      ← Main script
├── config.py               ← Your GitHub token (never share this)
├── config.example.py       ← Template for config.py
├── requirements.txt        ← Required Python packages
├── run_updater.bat         ← Double-click to run manually on Windows
├── .gitignore              ← Prevents config.py from being pushed to GitHub
└── ai_student_details.xlsx ← Your Excel file (not included in repo)
```

---

## 🔧 Troubleshooting

**Script says "Permission Denied" when saving:**
→ Close the Excel file first, then run the script again.

**HackerRank showing 0 stars for everyone:**
→ HackerRank may have changed their page structure. Open an issue on this repo.

**GitHub showing "Error" for many students:**
→ Your token may be invalid. Generate a new one and update `config.py`.

**"Module not found" error:**
→ Virtual environment is not active. Run `.venv\Scripts\activate` first.

---

## 📦 Libraries Used

| Library | Purpose |
|---|---|
| `openpyxl` | Read and write Excel files |
| `requests` | Make HTTP requests to APIs |
| `beautifulsoup4` | Parse HackerRank HTML pages |

All free and open source. Install with:
```bash
pip install openpyxl requests beautifulsoup4
```

---

## 📄 License

This project is open source and free to use.
