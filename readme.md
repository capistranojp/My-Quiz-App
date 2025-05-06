# 📚 Quiz App (Tkinter + CustomTkinter)

A simple, elegant desktop quiz application built using Python, `tkinter`, and `customtkinter`. Users can create, take, and track multiple quizzes with automatic score logging.

## ✨ Features

- 📝 Create custom quizzes with multiple-choice questions.
- 🎯 Take quizzes and get instant scores.
- 📊 View and manage score history.
- 📁 Data stored in `quizzes.xlsx` using `openpyxl`.

## 📦 Requirements

- Python 3.7+
- `customtkinter`
- `openpyxl`

Install dependencies:
```bash
pip install customtkinter openpyxl
```

## 🚀 Run the App

```bash
python quiz_app.py
```

## 📁 File Structure

```
├── quiz_app.py         # Main app script
├── quizzes.xlsx        # Auto-generated storage file
└── README.md
```

## 🧠 Notes

- Quiz data is saved in an Excel file (`quizzes.xlsx`) with separate sheets for each quiz and a `Scores` sheet for history.
- First-time run generates the file automatically.

## 📸 UI Preview

*(Add screenshot here if desired)*

---

Made with ❤️ using Python & CustomTkinter
