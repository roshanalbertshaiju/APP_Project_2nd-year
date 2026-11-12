# ExamOpt — Intelligent Examination Scheduling & Timetable Optimization System

[![Python Version](https://img.shields.io/badge/Python-3.10%20%7C%203.11%20%7C%203.12-blue.svg)](https://www.python.org/)
[![Database](https://img.shields.io/badge/Database-SQLite%203%20(3NF)-green.svg)](https://www.sqlite.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-Tkinter%20%2F%20ttk-orange.svg)](https://docs.python.org/3/library/tkinter.html)
[![Symbolic Engine](https://img.shields.io/badge/Engine-SymPy%201.14-informational.svg)](https://www.sympy.org/)
[![Test Suite](https://img.shields.io/badge/Tests-53%2F53%20Passed%20(100%25)-brightgreen.svg)]()
[![Repository](https://img.shields.io/badge/GitHub-roshanalbertshaiju%2FAPP__Project__2nd--year-blueviolet.svg)](https://github.com/roshanalbertshaiju/APP_Project_2nd-year)

---

### Academic Institution & Course Information
* **Institution:** SRM Institute of Science and Technology, Tiruchirappalli Campus
* **Department:** Department of Computer Science & Engineering
* **Course:** 21CSC203P — Advanced Programming Practice (APP)
* **Academic Year:** 2nd Year B.Tech (2026–2027)
* **GitHub Repository:** [https://github.com/roshanalbertshaiju/APP_Project_2nd-year](https://github.com/roshanalbertshaiju/APP_Project_2nd-year)

### Project Team Members
* **Roshan Albert Shaiju** (RA2311003010028)
* **Suhas V A**
* **Subaash G**

---

## 📋 Table of Contents
1. [Project Overview & Problem Statement](#-project-overview--problem-statement)
2. [Key Implemented Modules & Functionalities](#-key-implemented-modules--functionalities)
3. [System Architecture](#-system-architecture)
4. [Screenshots of Project Output](#-screenshots-of-project-output)
5. [Database Files & SQL Scripts](#-database-files--sql-scripts)
6. [Required Libraries & Dependencies](#-required-libraries--dependencies)
7. [Step-by-Step Execution Instructions](#-step-by-step-execution-instructions)
8. [Automated Verification & Test Suite](#-automated-verification--test-suite)
9. [Project Repository Layout](#-project-repository-layout)

---

## 🎯 Project Overview & Problem Statement

Academic examination scheduling is a classic **NP-hard combinatorial optimization problem**. In higher educational institutions like SRM IST, manual scheduling across dozens of departments and hundreds of enrolled students routinely causes:
* **Student Exam Clashes:** Students enrolled in multiple courses assigned to concurrent time slots.
* **Faculty Invigilation Overlaps:** Double-booking proctors across distant exam halls.
* **Room Capacity Violations & Inefficiencies:** Overcrowded halls or vast auditoriums assigned to tiny elective batches.
* **Consecutive Fatigue:** Scheduling rigorous exams back-to-back without restorative rest intervals.

**ExamOpt** resolves these challenges through an engineered, clean-architecture software solution implemented entirely in Python and SQLite. The platform couples a deterministic Minimum Remaining Values (MRV) heuristic scheduling engine with a strict 6-rule constraint validator, a meta-heuristic Genetic Algorithm (GA), a multi-core parallel candidate search (`multiprocessing`), symbolic mathematical objective modeling (SymPy), an Obsidian Dark modern GUI studio, and automated industrial PDF/Excel/CSV document exporters.

---

## 🚀 Key Implemented Modules & Functionalities

ExamOpt is engineered into decoupled, cohesive packages adhering to SOLID principles:

### 1. Relational Database Layer (`database/`)
* **3NF SQLite Relational Architecture:** Fully normalized database schema (`database/schema.sql`, `database/schema.py`) enforcing foreign key constraints (`PRAGMA foreign_keys = ON;`) with `ON DELETE CASCADE` and `ON DELETE SET NULL`.
* **Data Access Object (DAO) Pattern:** Thread-safe `Database` class encapsulating transactional CRUD operations with parameterized queries preventing SQL injection.
* **High-Speed Composite Indexes:** Pre-indexed foreign keys (`idx_enrollments_subject`, `idx_exams_slot`, `idx_exams_room`) enabling sub-millisecond table joins.
* **Database Dump & Seed Scripts:** Bundled SQL scripts (`schema.sql`, `seed_data.sql`, `examopt.sql`) and binary database (`examopt.db`).

### 2. Domain Entities & Object Models (`models/`)
* **Typed Data Models:** Clean Python data classes representing academic entities:
  * `Student`: Identification, departmental affiliation, semester, and registered subject codes.
  * `Subject`: Course code, title, department, semester, and duration (default 180 min).
  * `Faculty`: Invigilator record, department, and active availability toggle.
  * `Room`: Venue identifier, hall name, seating capacity threshold, and maintenance flag.
  * `TimeSlot`: Date, start time, end time, and active availability toggle.
  * `Exam`: Exam ID, subject, assigned slot, assigned room, assigned invigilator, student count, and status badge.
  * `Enrollment`: Composite associative entity mapping student-to-subject course registrations.

### 3. Service Facade Gateway (`services/`)
* **Unified Service Facade (`ExamService`):** Single point of contact for the presentation layers, coordinating database persistence, scheduling invocations, constraint audits, and analytical evaluations.

### 4. Intelligent Scheduling & Optimization Engines (`scheduling/`)
* **MRV Heuristic Scheduler (`scheduler.py`):** Schedules subjects using the *Minimum Remaining Values* heuristic (prioritizing subjects with largest student enrollments and most constrained slot options).
* **Best-Fit Room Allocator (`room_allocator.py`):** Computes minimal surplus capacity:
  $$\Delta_{\mathrm{capacity}} = C_{\mathrm{room}} - N_{\mathrm{students}} \quad (\text{where } \Delta_{\mathrm{capacity}} \ge 0)$$
  minimizing seat wastage while guaranteeing zero venue double-booking.
* **Meta-Heuristic Genetic Algorithm (`genetic_scheduler.py`):**
  * Evolutionary chromosome representation mapping exams to slot-room genes.
  * Fitness function combining inverse hard violations with weighted soft penalties.
  * Tournament selection, two-point crossover, and random perturbation mutation operators over 50 generations.
* **Multi-Core Parallel Search (`parallel_scheduler.py`):**
  * Utilizes `concurrent.futures.ProcessPoolExecutor` to bypass Python's Global Interpreter Lock (GIL).
  * Generates and evaluates candidate timetable variations in parallel across CPU logical cores, achieving a **3.19x speedup** on quad-core architectures.
* **Functional Pipeline Evaluator (`functional_evaluator.py`):**
  * Pure functional programming pipeline using `map()`, `filter()`, `functools.reduce()`, and list comprehensions.
  * Computes room utilization rates, detects consecutive student exam sessions, and calculates faculty workload variance without mutable global state.

### 5. Academic Constraint Validator (`validation/`)
Strict audit engine enforcing **6 Essential Academic Rules**:
1. **Rule 1 — Student Conflict Rule (Hard):** $\forall s \in \text{Students}, \text{count}(\text{Exams}(s, \text{slot})) \le 1$.
2. **Rule 2 — Room Capacity Rule (Hard):** $N_{\mathrm{students}} \le \text{Capacity}(\text{Room})$.
3. **Rule 3 — Room Availability Rule (Hard):** Venues marked unavailable or under renovation cannot be assigned.
4. **Rule 4 — Room Double-Booking Rule (Hard):** A physical hall can host at most one exam per time slot.
5. **Rule 5 — Faculty Availability & Duty Rule (Hard):** Proctors cannot be assigned to concurrent exams or when unavailable.
6. **Rule 6 — Time Slot Validity Rule (Hard):** Exams must fall within active, authorized institution scheduling windows.

### 6. SymPy Symbolic Mathematical Modeling (`optimization/`)
* Formulates the global timetable penalty loss analytically:
  $$\mathcal{L}_{\mathrm{total}} = M_{\mathrm{hard}} \cdot N_{\mathrm{hard}} + w_{\mathrm{consec}} \cdot P_{\mathrm{consec}} + w_{\mathrm{fac}} \cdot P_{\mathrm{fac}} + w_{\mathrm{waste}} \cdot P_{\mathrm{waste}}$$
* Computes analytical partial derivatives $\frac{\partial \mathcal{L}}{\partial w_i}$ for parameter sensitivity profiling and gradient evaluation.

### 7. Distributed TCP Socket Architecture (`networking/`)
* **Custom Binary Framing Protocol (`protocol.py`):** 4-byte big-endian length prefix prepending JSON-RPC payloads, guaranteeing stream framing over TCP sockets.
* **Daemon TCP Server (`exam_server.py`):** Non-blocking threaded server listening on `127.0.0.1:8080`, supporting remote requests (`PING`, `SCHEDULE`, `VALIDATE`, `SYMBOLIC_MODEL`).
* **Client SDK (`exam_client.py`):** High-level client API for programmatic network interaction.

### 8. Presentation Layer & Studio GUIs (`gui/`)
* **Modern Studio GUI (`modern_app.py`):** Dark-mode Obsidian-themed interface featuring KPI statistic cards, real-time timetable grids, conflict resolution badges, and interactive navigation tabs.
* **Classic Institutional GUI (`app.py`):** High-contrast standard Tkinter desktop layout.
* **Interactive CLI (`main.py`):** Command-line pipeline for headless servers and continuous integration.

### 9. Multi-Format Industrial Exporter (`utils/`)
* **CSV Exporter:** Standard comma-delimited tabular timetable.
* **OpenPyXL Excel Exporter:** Formatted `.xlsx` spreadsheet with styled headers, alternating row colors, auto-fitted columns, and verification badges.
* **ReportLab PDF Exporter:** Printable examination bulletin with SRM IST header, institutional metadata, and paginated timetable grid.

---

## 🏛️ System Architecture

ExamOpt follows a **5-Tier Clean Architecture**:

```
+-------------------------------------------------------------------------------+
|                           1. PRESENTATION LAYER                               |
|   Modern Studio GUI (Obsidian)  |  Classic Institutional GUI  |  Command CLI  |
+---------------------------------------+---------------------------------------+
                                        |
+---------------------------------------v---------------------------------------+
|                         2. SERVICE FACADE LAYER                               |
|          ExamService (Unified Orchestrator of Domain, DB & Validation)        |
+---------------------------------------+---------------------------------------+
                                        |
+---------------------------------------v---------------------------------------+
|                          3. DOMAIN ENGINES LAYER                              |
|   MRV Scheduler  *  Best-Fit Allocator  *  6-Rule Validator  *  GA  *  SymPy  |
+---------------------------------------+---------------------------------------+
                                        |
+---------------------------------------v---------------------------------------+
|                    4. PERSISTENCE & NETWORKING LAYER                          |
|    SQLite 3NF Database DAO (examopt.db)  |  TCP Socket JSON-RPC Daemon (:8080)|
+---------------------------------------+---------------------------------------+
                                        |
+---------------------------------------v---------------------------------------+
|                         5. OUTPUT DELIVERY LAYER                              |
|       Real-Time Grid  |  CSV Exporter  |  OpenPyXL Excel  |  ReportLab PDF    |
+-------------------------------------------------------------------------------+
```

---

## 📸 Screenshots of Project Output

All output screenshots are stored in the `screenshots/` directory:

### 1. System Architecture & Component Hierarchy
![System Architecture](screenshots/01_system_architecture.png)
*Figure 1: Five-tier clean architectural model depicting component boundaries and data flows.*

---

### 2. Modern GUI Timetable Studio
![Modern GUI Studio](screenshots/02_modern_gui_timetable_studio.png)
*Figure 2: Modern Obsidian Studio showing real-time timetable grid, KPI cards, and validation badges.*

---

### 3. Automated Academic Constraint Validator Diagnostics
![Constraint Validator](screenshots/03_constraint_validation_diagnostics.png)
*Figure 3: Six-rule academic integrity audit verifying zero hard constraint violations.*

---

### 4. Genetic Algorithm Optimization Convergence
![GA Convergence](screenshots/04_ga_optimization_convergence.png)
*Figure 4: Evolutionary convergence curve over 50 generations showing hard violation elimination.*

---

### 5. Multiprocessing Parallel Candidate Search Benchmark
![Parallel Benchmark](screenshots/05_parallel_multiprocessing_benchmark.png)
*Figure 5: Multi-core benchmark demonstrating 3.19x execution speedup across 4 CPU cores.*

---

### 6. Official Timetable PDF Export (SRM IST Branded)
![Official Timetable PDF](screenshots/06_official_timetable_pdf.png)
*Figure 6: Generated PDF timetable report exported via ReportLab with institutional layout.*

---

### 7. End-to-End Scheduling Workflow
![Scheduling Workflow](screenshots/07_scheduling_workflow.png)
*Figure 7: Seven-stage operational pipeline from SQLite database initialization to multi-format document delivery.*

---

### 8. Command-Line Execution & Milestone Review Checkpoints
![CLI Execution](screenshots/08_cli_execution_and_tests.png)
*Figure 8: Console output showing 10/10 academic milestone review scenarios passing.*

---

### 9. SymPy Symbolic Mathematical Modeling & Sensitivity Gradients
![SymPy Modeling](screenshots/09_sympy_symbolic_modeling.png)
*Figure 9: Formal loss objective function and partial derivative sensitivity analysis.*

---

### 10. Distributed TCP Socket Client-Server Network Architecture
![TCP Socket Protocol](screenshots/10_distributed_tcp_socket_network.png)
*Figure 10: Binary 4-byte length-prefixed JSON-RPC communication over port 8080.*

---

## 🗄️ Database Files & SQL Scripts

The repository includes complete, ready-to-use SQL scripts in the `database/` directory:

| File | Description | Purpose |
| :--- | :--- | :--- |
| **`database/schema.sql`** | Complete DDL Schema Script | Defines all 7 relational tables, primary keys, cascading foreign keys, and performance indexes. |
| **`database/seed_data.sql`** | Sample DML Seed Dataset | Inserts 6 subjects, 80 students across 3 departments, multi-course enrollments, 6 faculty members, 7 examination halls, and 6 time slots. |
| **`database/examopt.sql`** | All-in-One SQL Dump | Combines DDL schema and DML seed data with transaction wrappers (`BEGIN TRANSACTION ... COMMIT;`) for instant recreation. |
| **`examopt.db`** | Binary SQLite 3 Database | Pre-seeded database file for out-of-the-box execution without preliminary setup. |

### Loading Database via SQLite CLI:
```powershell
# Recreate the complete database from the SQL script
sqlite3 examopt.db < database/examopt.sql
```

---

## 📦 Required Libraries & Dependencies

All required dependencies are specified in `requirements.txt`:

```text
# Core Mathematical & Algorithmic Engines
sympy>=1.12.0
tabulate>=0.9.0

# Document & Spreadsheet Export Engines
openpyxl>=3.1.2
reportlab>=4.0.0
python-docx>=1.0.0
python-pptx>=1.0.0

# Data Visualization, Image Processing & Screen Renders
matplotlib>=3.7.0
pillow>=10.0.0
pymupdf>=2.0.0
```

### Dependency Installation:
From the root project directory:
```powershell
python -m pip install -r requirements.txt
```

---

## 💻 Step-by-Step Execution Instructions

### 1. Launching the Graphical User Interface (GUI)
Launch the default **Modern Obsidian Studio**:
```powershell
python main.py
```
*(Or double-click `run_modern_gui.bat` / `run_examopt.bat` on Windows)*

Launch the **Classic Institutional UI**:
```powershell
python main.py --theme classic
```
*(Or double-click `run_classic_gui.bat` on Windows)*

---

### 2. Running the End-to-End CLI Pipeline
Executes database initialization, sample data seeding, MRV timetable generation, 6-rule validation, formatted timetable printing, and 10 milestone scenarios:
```powershell
python main.py --cli
```

---

### 3. Running the Academic Review Scenarios
Runs the complete 10-milestone academic verification suite:
```powershell
python main.py --scenarios
```
*(Alias: `python main.py --verify-100`)*

---

### 4. Running Meta-Heuristic Genetic Algorithm Optimization
Executes 50 generations of evolutionary chromosome optimization:
```powershell
python main.py --ga
```

---

### 5. Multi-Core Parallel Scheduling & Benchmarking
Search candidate timetables across multiple CPU processes:
```powershell
python main.py --parallel
```

Benchmark sequential execution against multi-core parallel execution:
```powershell
python main.py --benchmark
```

---

### 6. Exporting Timetables (CSV, Excel, PDF)
Generates exported timetable files into the `exports/` folder:
```powershell
python main.py --export-all
```
Outputs created:
* `exports/schedule.csv` — Comma-delimited data
* `exports/schedule.xlsx` — Formatted OpenPyXL spreadsheet
* `exports/schedule.pdf` — Branded ReportLab printable PDF

---

### 7. Launching Distributed TCP Socket Server & Client
Start the socket daemon on port 8080:
```powershell
python main.py --server --port 8080
```

In a separate terminal, test client queries:
```python
from networking.exam_client import ExamSocketClient
client = ExamSocketClient(port=8080)
print(client.ping())
print(client.generate_schedule())
client.close()
```

---

## 🧪 Automated Verification & Test Suite

ExamOpt features **53 automated unit and integration tests** covering all modules:

```powershell
python -m unittest discover -s tests -v
```

### Test Coverage Summary:
* `test_crud.py` (6 tests) — CRUD operations across all 7 database entities.
* `test_exports.py` (3 tests) — CSV, Excel (`.xlsx`), and PDF generation integrity.
* `test_functional.py` (6 tests) — Functional pipeline (`map`, `filter`, `reduce`), room utilization, seat wastage.
* `test_genetic.py` (4 tests) — Population initialization, two-point crossover, mutation, full GA optimization.
* `test_multiprocessing.py` (2 tests) — Parallel schedule candidate search across CPU workers.
* `test_scheduler.py` (5 tests) — MRV heuristic scheduling, best-fit room allocation, double-booking prevention.
* `test_sockets.py` (4 tests) — TCP socket ping-pong, remote scheduling, remote validation, remote SymPy model.
* `test_sympy.py` (5 tests) — Symbolic expression generation, loss evaluation, partial derivatives, sensitivity analysis.
* `test_validator.py` (8 tests) — 6-rule constraint validator, detecting student clashes, capacity overflows, unavailable entities.
* `test_scenarios.py` (10 tests) — 10 academic milestone verification scenarios.

**Result: 53 / 53 Tests Passed (100% Pass Rate).**

---

## 📂 Project Repository Layout

```text
APP ExamOPS/
├── .gitignore                                 # Git ignore patterns for clean commits
├── README.md                                  # Comprehensive project documentation
├── requirements.txt                           # Third-party Python dependencies
├── config.py                                  # Global configuration & environment paths
├── main.py                                    # Unified CLI and GUI application entry point
├── examopt.db                                 # Pre-seeded SQLite 3 relational database
├── run_examopt.bat                            # Windows launcher (Default Modern GUI)
├── run_modern_gui.bat                         # Windows launcher (Modern Obsidian Studio)
├── run_classic_gui.bat                        # Windows launcher (Classic Institutional UI)
│
├── database/                                  # Database Layer
│   ├── __init__.py
│   ├── database.py                            # SQLite DAO implementation & CRUD methods
│   ├── schema.py                              # DDL schema strings & table definitions
│   ├── schema.sql                             # Pure ANSI/SQLite DDL SQL script
│   ├── seed_data.sql                          # DML seed data SQL script (80 students, 6 subjects)
│   └── examopt.sql                            # Complete all-in-one database SQL dump
│
├── models/                                    # Domain Entities
│   ├── __init__.py
│   ├── student.py                             # Student data class
│   ├── subject.py                             # Subject data class
│   ├── faculty.py                             # Faculty/Invigilator data class
│   ├── room.py                                # Room/Venue data class
│   ├── timeslot.py                            # TimeSlot data class
│   ├── exam.py                                # Scheduled Exam data class
│   └── enrollment.py                          # Course enrollment mapping
│
├── services/                                  # Service Facade Layer
│   ├── __init__.py
│   └── exam_service.py                        # Unified facade coordinating business logic
│
├── scheduling/                                # Timetable Engines & Optimization
│   ├── __init__.py
│   ├── scheduler.py                           # Heuristic MRV timetable scheduler
│   ├── room_allocator.py                      # Best-fit minimal-wastage room allocator
│   ├── genetic_scheduler.py                   # Meta-heuristic Genetic Algorithm (GA)
│   ├── parallel_scheduler.py                  # Multi-core candidate search (ProcessPool)
│   └── functional_evaluator.py                # Pure functional evaluation pipeline
│
├── validation/                                # Academic Constraint Validation
│   ├── __init__.py
│   └── validator.py                           # 6-Rule academic constraint validator
│
├── optimization/                              # Mathematical Modeling
│   ├── __init__.py
│   └── symbolic_model.py                      # SymPy symbolic loss model & gradients
│
├── networking/                                # Distributed Client-Server Protocol
│   ├── __init__.py
│   ├── protocol.py                            # 4-byte length-prefixed framing protocol
│   ├── exam_server.py                         # Multithreaded TCP socket server (daemon)
│   └── exam_client.py                         # TCP socket client API
│
├── gui/                                       # Graphical User Interface
│   ├── __init__.py
│   ├── modern_app.py                          # Modern Obsidian Dark Studio GUI
│   ├── app.py                                 # Classic Institutional Tkinter GUI
│   └── views/                                 # Modular UI View Panels
│       ├── __init__.py
│       ├── timetable_view.py                  # Timetable grid & generation view
│       ├── student_view.py                    # Student management view
│       ├── subject_view.py                    # Subject management view
│       ├── faculty_view.py                    # Faculty management view
│       ├── room_view.py                       # Room management view
│       ├── timeslot_view.py                   # TimeSlot management view
│       ├── genetic_view.py                    # Genetic Algorithm controls & graph
│       ├── benchmark_view.py                  # Parallel CPU benchmark view
│       ├── network_view.py                    # TCP socket server/client view
│       ├── verification_view.py               # Constraint audit view
│       └── viva_suite_view.py                 # Academic evaluation suite view
│
├── utils/                                     # Utilities & Exporters
│   ├── __init__.py
│   ├── export_manager.py                      # CSV, OpenPyXL (.xlsx), and PDF exporters
│   ├── sample_data.py                         # SRM IST realistic data generator
│   └── helpers.py                             # General helper functions
│
├── tests/                                     # Automated Test Suite (53 Tests)
│   ├── __init__.py
│   ├── test_crud.py                           # Database CRUD unit tests
│   ├── test_exports.py                        # Export integrity unit tests
│   ├── test_functional.py                     # Functional evaluation tests
│   ├── test_genetic.py                        # Genetic Algorithm unit tests
│   ├── test_multiprocessing.py                # ProcessPool parallel search tests
│   ├── test_scheduler.py                      # Scheduler & allocator unit tests
│   ├── test_sockets.py                        # TCP socket communication tests
│   ├── test_sympy.py                          # SymPy symbolic math tests
│   ├── test_validator.py                      # 6-rule constraint validator tests
│   └── test_scenarios.py                      # 10 milestone review scenario tests
│
├── screenshots/                               # Output Screenshots for Documentation
│   ├── 01_system_architecture.png             # 5-tier clean architectural model
│   ├── 02_modern_gui_timetable_studio.png     # Modern Obsidian Dark UI studio
│   ├── 03_constraint_validation_diagnostics.png # 6-rule validator diagnostics
│   ├── 04_ga_optimization_convergence.png     # GA fitness convergence chart
│   ├── 05_parallel_multiprocessing_benchmark.png # CPU multi-core speedup benchmark
│   ├── 06_official_timetable_pdf.png          # Official SRM IST timetable PDF
│   ├── 07_scheduling_workflow.png             # End-to-end scheduling workflow
│   ├── 08_cli_execution_and_tests.png         # CLI execution & 10/10 scenarios passed
│   ├── 09_sympy_symbolic_modeling.png         # SymPy loss equations & gradients
│   └── 10_distributed_tcp_socket_network.png  # TCP socket network architecture
│
└── exports/                                   # Generated Export Artifacts
    ├── schedule.csv                           # Exported CSV timetable
    ├── schedule.xlsx                          # Exported OpenPyXL styled Excel workbook
    └── schedule.pdf                           # Exported ReportLab printable PDF
```

---

## 📜 Academic Integrity & License
This project was designed and developed for the **21CSC203P Advanced Programming Practice (APP)** course at **SRM Institute of Science and Technology, Tiruchirappalli Campus**.
Licensed under the [MIT License](LICENSE).
