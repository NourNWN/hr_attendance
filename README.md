# HR Attendance Automation Tool 🚀

An elegant, production-ready desktop application built with Python and Tkinter to automate monthly HR attendance calculations from fingerprint Excel data.

## ✨ Features
- **Symmetric & Clean GUI**: Built with a modern teal-themed user interface.
- **Smart Validation**: Input fields are guarded against wrong data types (accepts only time/numeric formats).
- **Interactive Tooltips**: Built-in guidance for every setting to optimize user experience.
- **Dynamic Workflow**: Automatically triggers and opens the generated review file for HR verification.
- **Engineered Logic**: Handles complex calculations for late minutes (with grace periods), early departures, overtime, and custom leaves (AM/PM).

## 🛠️ Tech Stack
- **Language**: Python 3.11
- **GUI Framework**: Tkinter
- **Data Processing**: Pandas, OpenPyXL
- **Packaging**: PyInstaller (Compiled into an independent `.exe`)

## 📦 Project Structure
- `GUI.py`: Handles styling, user interactions, tooltips, and validations.
- `hr_attendance.py`: Core mathematical engine for checking punch-in/out records.
- `helpers.py` / `settings_input.py`: Modular code helper components.

## 👨‍💻 Author
**Eng. Nour Nasser** - *Designed & Engineered with passion.*