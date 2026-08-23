# DS Project 2 — Hostel Student Management System

A small C++ data structures project. It keeps hostel student records and lets you
add, search, delete, undo a delete, and sort them.

There are two ways to use it: a **console menu** and a **web page** (HTML + CSS)
served on localhost by a tiny C++ web server.

## Data structures used

| Structure | Where it is used | File |
|---|---|---|
| **Array** | `students[100]` holds all the student records. Insert adds at the end, delete shifts the elements left, search is a linear search. | `src/hostel.h` |
| **Stack** | `stack_[100]` with `top_`. Every deleted student is **pushed** on it, and *Undo* **pops** the last one back into the array. | `src/hostel.h` |
| **String** | `std::string` for name and course — used for case-insensitive search (`contains`) and for sorting by name (bubble sort comparing strings). | `src/hostel.h` |

## Files

```
ds_project_2/
├── src/
│   ├── hostel.h       all the data structure code (array, stack, string)
│   ├── console.cpp    menu driven console version
│   └── server.cpp     small web server for the HTML page
├── web/
│   ├── index.html     the UI (plain HTML, no JavaScript)
│   └── style.css      the styling
├── docs/
│   ├── DS_Project_2_Report.pdf         full project report (for submission)
│   ├── DS_Project_2_Presentation.pptx  16-slide presentation
│   ├── report.html                     source of the report
│   └── make_ppt.py                     script that builds the presentation
├── Makefile
└── README.md
```

All the logic lives in one small header, `src/hostel.h`, so it is easy to read
and easy to reuse in a bigger project later.

## Project report

A complete write-up is in **[docs/DS_Project_2_Report.pdf](docs/DS_Project_2_Report.pdf)** —
objective, why each data structure was chosen, code walkthrough, how the web UI works,
test results, time complexity, limitations, and a set of likely viva questions with answers.

Regenerate it after editing `docs/report.html` with `make report`.

A 16-slide presentation is in **[docs/DS_Project_2_Presentation.pptx](docs/DS_Project_2_Presentation.pptx)**,
built by `docs/make_ppt.py` (`make ppt`, needs `python-pptx`).

## How to run

Needs `g++` and Linux/macOS (the server uses POSIX sockets).

### Web version

```bash
make web
```

Then open <http://localhost:8080> in your browser.

> Run it from the project folder, because the server reads `web/index.html`
> and `web/style.css` using relative paths.

### Console version

```bash
make cli
```

## What the web page does

* **Add Student** — inserts into the array
* **Search** — case-insensitive string search on name, course or roll number
* **Sort by name** — bubble sort comparing strings
* **delete** — removes from the array and pushes the record on the stack
* **Deleted Students (Stack)** — shows the stack with the top marked
* **Undo last delete (pop)** — pops the stack and puts the student back

The UI is only HTML and CSS. There is no JavaScript — every button is a normal
HTML form or link, and the C++ server builds the page by filling in the
placeholders (`{{ROWS}}`, `{{STACK}}`, `{{COUNT}}`, `{{SEARCH}}`) in
`web/index.html`.

## Notes / limitations

* Data is kept in memory only, so it is lost when the program stops.
* Maximum 100 students (`MAX` in `src/hostel.h`).
* The server handles one request at a time — it is meant for learning, not for
  real use.
