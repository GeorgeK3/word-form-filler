# Word Form Filler (Python)

A simple **Python GUI application** that fills a **Word (.docx) template** using user input.

The program allows users to enter personal information through a form and automatically inserts the values into a Word document.

The filled document is then saved with the user's **last name as the filename**.

---

## Features

* Simple **Tkinter GUI**
* Automatic **Word document editing**
* Input **validation**
* Field highlighting for invalid input
* Keyboard navigation with **arrow keys**
* Save output as a new `.docx` file

---

## Input Validation

The program checks that:

* **Name / Last Name** → letters only
* **Age** → number between 1–120
* **Car ID** → format `AAA-1234`


## Usage

Run the application:

```
python form_filler.py
```

1. Fill the form fields
2. Press **Enter** or click **Fill & Save**
3. A new `.docx` file will be created using the last name as the filename

---

## Technologies Used

* Python
* Tkinter
* python-docx

---

## Example Use Case

This project can be used for:

* Document automation
* Office forms
* Data entry tools
* Learning Python GUI development

