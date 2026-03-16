# WorkElate AI-Native Office Intern Assignment

## Project Overview

This project is developed as part of the **WorkElate Internship Evaluation Assignment**.
The application simulates an Excel-like spreadsheet interface where users can perform sorting, filtering, clipboard operations, and persistent data storage.

The goal of this implementation is to demonstrate **problem-solving ability, product thinking, clean architecture, and handling of real-world edge cases.**

---

## Tech Stack

* React (Frontend UI)
* Vite (Build Tool)
* JavaScript (Core Logic)
* CSS (Styling)
* Browser Clipboard API
* LocalStorage API

---

## Features Implemented

### Task 1 — Column Sort & Filter

* Column sorting cycle: **Ascending → Descending → None**
* Sorting works on **computed formula values**
* Excel-style **filter dropdown in column headers**
* Filtering hides rows without deleting original data
* Sorting & filtering implemented at **view layer only**
* Formulas always reference **original cell coordinates**

---

### Task 2 — Multi-Cell Copy & Paste

* Supports **Ctrl + V** paste from Excel / Google Sheets
* Handles **multi-row and multi-column tab-separated data**
* Paste actions are **undoable using Ctrl + Z**
* **Ctrl + C copies computed values**
* Supports **internal spreadsheet copy-paste**

---

### Task 3 — Local Storage Persistence

* Spreadsheet state **auto-saved with 500ms debounce**
* Restores:

  * Cell values
  * Formulas
  * Styles
  * Grid dimensions
* Undo / redo history **not persisted**
* Handles:

  * Storage quota limits
  * Corrupted local storage data safely

---

## Architecture & Approach

* A **central grid state model** manages spreadsheet data.
* Sorting & filtering are implemented as **derived view transformations**.
* Formula evaluation is handled through a **lightweight computation engine**.
* Clipboard integration uses the **native browser Clipboard API**.
* Persistence implemented via **debounced localStorage synchronization**.
* Undo / redo handled using **state history stacks**.

---

## Product & UX Decisions

* Column header click cycles sorting states for intuitive interaction.
* Filter dropdown shows **unique column values** similar to Excel.
* Multi-cell selection includes **visual highlighting**.
* Paste behaviour mimics **real spreadsheet experience**.
* Auto-save designed to be **non-blocking and smooth**.

---

## Edge Cases Handled

* Sorting columns containing formula-derived values.
* Pasting large datasets from external spreadsheet tools.
* Undo functionality after bulk paste operations.
* Safe recovery from **corrupted saved state**.
* Graceful handling of **local storage size limits**.

---

## Setup Instructions

1. Clone the repository:

```
git clone <your-repo-link>
```

2. Install dependencies:

```
npm install
```

3. Run development server:

```
npm run dev
```

---

## Demo Walkthrough

Loom Video Explanation:
(https://drive.google.com/file/d/1LJWb75pwb2Sxg212WZpPijBsQ6ZF40sA/view?usp=sharing)

---

## Conclusion

This implementation focuses on building a **functional, scalable, and user-friendly spreadsheet experience**, reflecting practical product engineering decisions and robust frontend architecture.
