## User Manual for Time-table Automation System

### 1. Overview

The Time-table Automation System is a command-line Python application designed to automate the scheduling of courses for academic sessions at a university. It processes course data, assigns rooms, schedules lectures, tutorials, and labs while avoiding conflicts (e.g., professor and room conflicts), and accounts for break times (morning and lunch breaks). The system generates timetables for each department-semester combination and outputs them to an Excel file (`timetables.xlsx`). It also identifies and attempts to schedule unscheduled courses, saving any remaining unscheduled courses to `unscheduled_courses.xlsx`.

This manual provides instructions for downloading, setting up, and using the system, along with usage scenarios, requirements satisfied, future work, and FAQs.

---

### 2. Instructions for Downloading the Software from GitHub

#### Accessing the Repository

- Visit the GitHub repository: `https://github.com/<your-team-username>/timetable-automation` (replace `<your-team-username>` with your team’s GitHub username).
- If the repository is private, request access from the project team.

#### Cloning the Repository

- Open a terminal (Command Prompt on Windows, Terminal on macOS/Linux) and run: git clone [remote repo link]()

- This downloads the project files to a `timetable-automation` folder on your machine.

#### Downloading as ZIP

- Alternatively, on the GitHub repository page, click the green “Code” button and select “Download ZIP”.
- Extract the ZIP file to a folder on your machine.

---

### 3. How to Set Up the Software

#### Prerequisites

- **Python 3.9 or later**: Download from [python.org](https://www.python.org/downloads/).
- **Git**: Download from [git-scm.com](https://git-scm.com/downloads) (optional, for cloning).
- **pip**: Python’s package manager (included with Python).

#### Setup Steps

1. **Navigate to the Project Directory**:

- Open a terminal and navigate to the project folder:
  ```
  cd timetable-automation
  ```

2. **Create a Virtual Environment**:

- Create a virtual environment to manage dependencies:
  ```
  python -m venv venv
  ```
- Activate the virtual environment:
  - On Windows:
    ```
    venv\Scripts\activate
    ```
  - On macOS/Linux:
    ```
    source venv/bin/activate
    ```

3. **Install Dependencies**:

- The system requires the following Python libraries: `pandas`, `openpyxl`, and their dependencies. Install them using:
  ```
  pip install pandas openpyxl
  ```
- Alternatively, if a `requirements.txt` file is provided, run:
  ```
  pip install -r requirements.txt
  ```

4. **Place Configuration Files**:

- Ensure the required input files (`combined.xlsx` and `Rooms.csv`) are in the project directory (see Section 4 for details).

5. **Run the Application**:

- Run the script to generate timetables:
  ```
  python script.py
  ```
- The script will generate `timetables.xlsx` and, if applicable, `unscheduled_courses.xlsx`.
- **Screenshot Placeholder**: [Insert screenshot of the terminal showing the `python script.py` command and the output messages]

---

### 4. Setting Up Configuration Files

The system requires two input files to operate: `combined.xlsx` and `Rooms.csv`. These files must be placed in the project directory.

#### Required Configuration Files

1. **combined.xlsx**:

- **Purpose**: Contains course data for scheduling.
- **Format** (example):
  ```
  Course Code,Course Name,Department,Semester,Faculty,L,T,P,S,Elective
  CS301,Software Engineering,CSE,5,Dr. Smith,3,1,0,0,No
  EC201,Circuits,ECE,3,Dr. Jones,2,1,2,0,No
  ```
- **Fields**:
  - `Course Code`: Unique course identifier (e.g., CS301, b1CS101 for electives).
  - `Course Name`: Name of the course.
  - `Department`: Department (e.g., CSE, ECE, DSAI).
  - `Semester`: Semester number (e.g., 1 to 8).
  - `Faculty`: Instructor name.
  - `L,T,P,S`: Lecture, Tutorial, Practical, Self-study hours (integers; S is optional).
  - `Elective`: Yes/No, indicating if the course is an elective.

2. **Rooms.csv**:

- **Purpose**: Contains room data for scheduling.
- **Format** (example):
  ```
  roomNumber,type
  A101,LECTURE_ROOM
  LabA,COMPUTER_LAB
  Room201,SEATER_120
  ```
- **Fields**:
  - `roomNumber`: Room identifier (e.g., A101, LabA).
  - `type`: Room type (LECTURE_ROOM, COMPUTER_LAB, SEATER_120).

#### Steps to Configure

1. Place `combined.xlsx` and `Rooms.csv` in the project directory (same folder as the script).
2. Edit the files using Excel or a text editor to match your institution’s data.

- **Screenshot Placeholder**: [Insert screenshot of the project directory showing `combined.xlsx` and `Rooms.csv`]

---

### 5. Usage Scenarios

Since the system is a command-line application, usage involves running the script and viewing the output Excel files.

#### Scenario 1: Generate Timetables for All Departments and Semesters

1. **Prepare Input Files**:

- Ensure `combined.xlsx` and `Rooms.csv` are in the project directory with the correct data.
- **Screenshot Placeholder**: [Insert screenshot of the project directory with the input files]

2. **Run the Script**:

- Open a terminal, navigate to the project directory, and run:
  ```
  python script.py
  ```
- The script will:
  - Read course and room data.
  - Schedule lectures, tutorials, and labs while avoiding conflicts.
  - Allocate break times (morning break: 10:30-10:45; lunch break: dynamically between 12:30-14:30).
  - Generate `timetables.xlsx` with separate sheets for each department-semester combination (e.g., `CSE_5`, `ECE_3`).
- **Screenshot Placeholder**: [Insert screenshot of the terminal showing the script execution and completion message]

3. **View the Timetable**:

- Open `timetables.xlsx` in Excel.
- Each sheet shows a timetable with:
  - Days (Monday to Friday) as rows.
  - Time slots (9:00 to 18:30, in 30-minute intervals) as columns.
  - Course details (code, type, faculty, room) in each slot.
  - A legend at the bottom listing courses, faculty, and LTPS details.
- **Screenshot Placeholder**: [Insert screenshot of a timetable sheet in `timetables.xlsx`]

4. **Check for Unscheduled Courses**:

- If any courses couldn’t be scheduled, the script generates `unscheduled_courses.xlsx`.
- Open this file to view details of unscheduled courses (code, name, required vs. scheduled LTPS hours).
- **Screenshot Placeholder**: [Insert screenshot of `unscheduled_courses.xlsx` if generated]

---

### 6. Requirements Satisfied by the Current Version

The current version (as of April 23, 2025) satisfies the following requirements from the Excel sheet:

- **REQ-02-Config**: The system reads course data and room assignments from `combined.xlsx` and `Rooms.csv`.
- **REQ-03**: Courses are scheduled in classrooms based on department-semester combinations, with students split into sections (e.g., CSE_5 uses a dedicated lecture room).
- **REQ-04-CONFLICTS**: The system avoids direct time conflicts for professors and rooms. It schedules labs separately from lectures/tutorials but does not yet ensure labs are scheduled _after_ lectures/tutorials.
- **REQ-05**: Courses with the same name but different departments/semesters are scheduled separately (handled by grouping into department-semester timetables).
- **REQ-06**: Scheduling adheres to the LTPSC structure (e.g., 1.5 hours lecture = 3 slots, 2 hours lab = 4 slots, 1 hour tutorial = 2 slots).
- **REQ-09-BREAKS**: Morning break (10:30-10:45) and a dynamic 1-hour lunch break (between 12:30-14:30) are included in the schedule.
- **REQ-10-FACULTY**: Professors are not scheduled for overlapping classes (direct conflicts are avoided), but the 3-hour gap constraint is not fully enforced.
- **REQ-14-VIEW**: Timetables are exported to Excel (`timetables.xlsx`) for viewing by coordinators, faculty, and students.

---

### 7. Future Work

The following features are planned for future versions:

- **UI Development**: Add a user interface (e.g., Flask web app or Tkinter desktop app) to allow coordinators, faculty, and students to interact with the system directly (REQ-14-VIEW).
- **Exam Scheduling (REQ-15-EXAM)**: Implement exam timetable scheduling with seating arrangements and minimize exam days.
- **Analytics Reports (REQ-16-ANALYTICS)**: Add reports on room usage, instructor effort, and student effort using Matplotlib/Seaborn.
- **Faculty Preferences (REQ-11-FACULTY)**: Incorporate faculty scheduling preferences (e.g., preferred days/times).
- **Assistant Allocation (REQ-17-ASSIST)**: Allocate teaching/lab assistants for large courses (e.g., over 100 students).
- **Lunch Break Staggering (REQ-18-LUNCH)**: Stagger lunch breaks by department/semester to avoid overcrowding.
- **Google Calendar Integration (REQ-13-GCALENDER)**: Allow faculty and students to sync timetables with Google Calendar.
- **Dynamic Modifications (REQ-01)**: Support modifying existing timetables with minimal changes.
- **Elective Grouping (REQ-05)**: Fully implement elective grouping (e.g., b1, b2) to avoid overlaps unless allowed.
- **Professor Constraints (REQ-10-FACULTY)**: Fully enforce the 3-hour gap between consecutive classes for professors.
- **Lab Scheduling Order (REQ-04-CONFLICTS)**: Ensure labs are scheduled after lectures/tutorials for the same course.

---

### 8. FAQs

**Q: What should I do if the script fails to run?**

- A: Ensure `combined.xlsx` and `Rooms.csv` are in the project directory with the correct format. Check that all dependencies (`pandas`, `openpyxl`) are installed. Review the terminal error message for details.

**Q: What if some courses are unscheduled?**

- A: The script generates `unscheduled_courses.xlsx` with details of unscheduled courses. You can adjust the input data (e.g., reduce conflicts, add more rooms) and rerun the script.

**Q: Can I customize the time slots or break times?**

- A: Yes, but you’ll need to modify the constants in the code (`START_TIME`, `END_TIME`, morning break, lunch break range). A future version will include a configuration file for these settings.

**Q: How do I view the timetable for a specific department?**

- A: Open `timetables.xlsx` in Excel. Each sheet is named by department and semester (e.g., `CSE_5`).

**Q: Can faculty or students view their schedules directly?**

- A: Not yet. Currently, the timetable is exported to Excel for manual viewing. A UI for direct viewing will be added in the future.

**Q: What happens if a room type is missing in `Rooms.csv`?**

- A: The script will print a warning (e.g., “No LECTURE_ROOM type rooms found”) and use default placeholders (e.g., “No Lecture Room”). Ensure all room types are included in `Rooms.csv`.

---

### 9. Conclusion

This user manual provides a guide to setting up and using the Time-table Automation System, a command-line tool for generating academic timetables. The current version automates scheduling with conflict avoidance and break times, outputting timetables to Excel. Future versions will add a user interface, additional features, and enhanced scheduling constraints. For support, contact the development team.
