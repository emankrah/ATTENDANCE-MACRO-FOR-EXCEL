# ATTENDANCE-MACRO-FOR-EXCEL (Excel VBA)
Simple attendance taking program
 
A single VBA macro that automates weekly attendance entry in Excel — no manual cell editing, no mouse clicking through rows. Run the macro, answer a few prompts per employee, done.
 
## What it does
 
- Reads the employee list dynamically from an `EMPLOYEES` sheet (no hardcoded names — add or remove employees anytime, no code changes needed)
- Prompts for an attendance date, then walks through each employee asking for **Time In** and **Time Out**
- Validates every date and time entry, re-prompting on invalid input
- Skips absent employees automatically (blank Time In = no record)
- Allows a blank Time Out (for employees who forgot to clock out)
- Appends new records to the `ATTENDANCE` sheet without ever overwriting previous entries
- Separates each date's records with two blank rows
- Formats all times as `h:mm AM/PM` automatically
- Jumps back to the `ATTENDANCE` sheet when finished, so you land where the results are
## Workbook structure
 
### `EMPLOYEES` sheet
 
| NAME | STAFF ID |
|---|---|
| JOHN DOE| 1234|
| JOHNDOE 2 | 12345|
| ... | ... |
 
Add a row for a new hire, delete a row when someone leaves. That's it — the macro always reads this sheet fresh.
 
### `ATTENDANCE` sheet
 
| DATE | NAME | STAFF ID | TIME IN | TIME OUT |
|---|---|---|---|---|
 
The macro creates this header automatically if the sheet is empty. New records are appended below whatever is already there.
 
## Installation
 
1. Open your workbook and press `Alt + F11` to open the VBA editor.
2. `File → Import File...` and select `RecordAttendance.bas` from this repo.
   *(Or manually: `Insert → Module`, then paste the code from `src/RecordAttendance.bas`.)*
3. Make sure your workbook has two sheets named `EMPLOYEES` and `ATTENDANCE`, matching the structure above. (Sheet name matching is case-insensitive and ignores stray spaces.)
4. Save the file as **Excel Macro-Enabled Workbook (`.xlsm`)** — a regular `.xlsx` cannot store macros.
## Running it
 
- Press `Alt + F8`, select `RecordAttendance`, click **Run**.
- Or assign the macro to a button/shortcut for one-click access.
### Flow
 
1. Enter the attendance date when prompted (e.g. `20-04-26`).
2. For each employee, in order, enter:
   - **Time In** — leave blank to mark them absent (no record is created, and you move to the next person)
   - **Time Out** — leave blank if unknown (record is still saved, just with an empty Time Out)
3. After the last employee, choose whether to enter attendance for another date (repeats the whole flow) or finish.
## Time input rules
 
Accepted formats: `7:30`, `7:30 AM`, `07:30`, `18:30`, `6:00 PM`, etc.
 
| Field | If AM/PM is omitted... |
|---|---|
| Time In | Assumed **AM** (e.g. `7:30` → `7:30 AM`; `18:30` → `6:30 PM`, since it's already 24-hour) |
| Time Out | Assumed **PM** for hours 1–11 (e.g. `6:00` → `6:00 PM`; `18:00` → `6:00 PM`) |
 
Invalid entries (`banana`, `99:99`, `25:80`, etc.) are rejected with an error message, and the prompt repeats until a valid value is entered or the field is left blank.
 
## A note on macros and security
 
Macro-enabled files (`.xlsm`) are often flagged or blocked by company IT policies, especially if downloaded from the internet or emailed as attachments. If you're deploying this at work:
- Check with your IT team about macro policy before rolling it out
- Files may need to be unblocked (right-click → Properties → *Unblock*) or placed in an IT-approved Trusted Location
- Some organizations require macros to be digitally signed before they'll run
## Repo structure
