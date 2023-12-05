# INFO:
#  -> Student number must start with "06", be 10 characters long, and a valid number
#  -> After 15 minutes (MAX_OUT_TIME), the student will be logged back in with "NOR"
#  -> ADMIN_CODE -> Admin menu

# 31 red
# 32 green
# 33 yellow
# 34 blue
# 35 purple
# 36 cyan

import math
import time
from datetime import datetime, timedelta
import os
import json
import dateutil.parser
from pytimedinput import timedInput
from csv_parser import searchCSV, readCSV
import survey

# Constants
MAX_OUT_TIME = 15  # Minutes

USER_CSV = "prod(9-26-23).csv"
PRINT_ENABLED = True
ADMIN_CODE = "000"

# Globals
# {user, name, date_left}
out_students = []

# {
#   "user": [{user, first, last, grade}] (sorted by 06)
#   "last": [{user, first, last, grade}] (sorted by last name)
# }
users = readCSV(USER_CSV)
curr_mode = "normal"  # normal | admin | modes | csv_view
page = 0


# https://stackoverflow.com/questions/2520893/how-to-flush-the-input-stream
def flush_input():
    import msvcrt
    while msvcrt.kbhit():
        msvcrt.getch()


def log_input(prefix, student, current_datetime):
    """
    Writes data to the log.txt file in "append" (a) mode

    Format:
        PREFIX: DATE_TIME | Student: 06xxxxxxxx
    """
    with open("log.txt", "a") as log_file:
        log_file.write(f"{prefix}: {current_datetime} | Student: {student}\n")


def get_day_of_week():
    days = ["Monday", "Tuesday", "Wednesday",
            "Thursday", "Friday", "Sat.", "Sun."]
    return days[datetime.now().weekday()]


def loadOutStudents():
    """
    Reads out.txt, and reloads the data into memory
    """
    try:
        global out_students
        with open("out.txt", "r") as out_file:
            # Parses json file
            out_students = json.load(out_file)
    except:
        None


def check_students():
    """
    Checks if any students have run out of time, in comparison to when they left. 
    If so, mark them as NOR
    """
    for i, student in enumerate(out_students, start=1):
        # Gets current time as parsed date object
        dt = dateutil.parser.parse(str(student["date_left"]))
        if (datetime.now() - dt).total_seconds() > MAX_OUT_TIME * 60:
            print(f"Out of time: {student['user']}")

            # Get the current date and time
            offset_time = dt + timedelta(minutes=MAX_OUT_TIME)
            formatted_datetime = offset_time.strftime("%m/%d/%Y @ %H:%M:%S")
            log_input("NOR", student['user'] + " - " +
                      student["user"], formatted_datetime)

            # Removes student from out_students, and writes to file
            out_students.pop(i-1)
            with open("out.txt", "w+") as out_file:
                json.dump(out_students, out_file, default=str)


# Modes
def mode_normal():
    # Display count of students out of room "Currently out: 1 student" (yellow)
    if (len(out_students) > 0):
        print(
            f"\033[0;33mCurrently out\033[0m: {len(out_students)} student{'s' if len(out_students) > 1 else ''}\n")

    flush_input()
    num, timedOut = timedInput("Scan or Enter your Student Number: ", timeout=10,
                               resetOnInput=True, maxLength=10, allowCharacters="0123456789")

    # Check if any students have exceeded time limit, and assume they forgot to sign back in
    check_students()

    os.system("cls")

    # Display count of students out of room (yellow) "Currently out: 1 student"
    if (len(out_students) > 0):
        print(
            f"\033[0;33mCurrently out\033[0m: {len(out_students)} student{'s' if len(out_students) > 1 else ''}\n")

    # Re-render the input
    print(f"Scan or Enter your Student Number: {num}\n")

    # Mode switcher
    global curr_mode
    if num == ADMIN_CODE:
        curr_mode = "modes"
        return 0

    if timedOut:
        return 0

    # Check if the input is a valid integer and valid 06
    if num.isdigit() and len(num) == 10 and num.startswith("06"):
        # Get the current date and time
        current_datetime = datetime.now()
        formatted_datetime = current_datetime.strftime("%m/%d/%Y @ %H:%M:%S")

        # Find out_student
        idx = -1
        for i, student in enumerate(out_students):
            if student['user'] == num:
                idx = i
                break

        # Retrieve their data from the csv file
        userIdx = searchCSV(users["user"], int("1" + num))
        firstName = num
        lastName = ""
        if userIdx >= 0:
            firstName = users["user"][userIdx]["first"]
            lastName = users["user"][userIdx]["last"]

        if idx > -1:
            # Student returned
            out_students.pop(idx)

            log_input(
                " IN",  f"{num} - {firstName} {lastName}", formatted_datetime)

            # "Welcome back {NAME (first)}" (green)
            print(f"\033[0;32mWelcome back {firstName}!\033[0m")
        else:
            # Student left
            out_students.append({
                "user": num,
                "date_left": current_datetime,
                "first": firstName,
                "last": lastName
            })

            # log the number
            log_input(
                "OUT", f"{num} - {firstName} {lastName}", formatted_datetime)

            # Prepare the text to print
            text_to_print = f"                CCHS HALL PASS\r\n\n"
            text_to_print += f"Student:\n ->  {firstName + ' ' + lastName}\n ->  {num}\r\n\n"
            text_to_print += f"Printed at:\n ->  {get_day_of_week()} {formatted_datetime}\n\n\n"
            text_to_print += f"FROM: Rm. 147\n"
            text_to_print += f"TO: RR / Ft. Office / Guidance\n\n"
            text_to_print += f"Educators Signature:________________________\n\n\n"
            text_to_print += f">>>-----< DONT FORGET TO SIGN BACK IN >-----<<<\n\n"

            # Add the paper cut command 'GS V 0'
            text_to_print += chr(29) + 'V' + chr(1)

            if PRINT_ENABLED:
                import win32print
                import win32ui
                # Get the default printer
                printer_name = win32print.GetDefaultPrinter()

                # Open the printer
                hPrinter = win32print.OpenPrinter(printer_name)

                # Start a print job
                try:
                    hJob = win32print.StartDocPrinter(
                        hPrinter, 1, ("print_job", None, "RAW"))
                    try:
                        # Write the text to the printer
                        win32print.WritePrinter(
                            hPrinter, text_to_print.encode())
                    finally:
                        # End the print job
                        win32print.EndDocPrinter(hPrinter)
                finally:
                    # Close the printer
                    win32print.ClosePrinter(hPrinter)
            else:
                print("\n================================================\n\n" +
                      text_to_print + "\n================================================\n")

            # "Don't forget to sign/scan back in!" (yellow)
            print("\033[0;33mDon't forget to sign/scan back in!\033[0m\n")

        # Save currently out students as serialized json, incase server crashes or closes
        with open("out.txt", "w+") as out_file:
            json.dump(out_students, out_file, default=str)

        return 2
    else:
        print("Invalid Student Number. (06xxxxxxxx)\n")


def mode_admin():
    # Current terminal height - 17; sizing that doesnt include surrounding text
    PAGE_LENGTH = os.get_terminal_size().lines - 17

    global page

    if not os.path.isfile("log.txt"):
        open("log.txt", "x").close()  # Quick create file (close buffer)
    with open("log.txt", "r") as log_file:
        text = log_file.readlines()
        text.reverse()

        # Pagination and fancy colors
        MAX_PAGES = math.ceil(len(text)/PAGE_LENGTH)
        stylizedDate = datetime.now().strftime(
            '\033[30m\033[47m%m/%d/%Y\033[0m')
        print(
            f"Admin Log ({math.floor(page/PAGE_LENGTH) + 1}/{MAX_PAGES}):\n")
        print(
            "".join(text[slice((page * PAGE_LENGTH), ((page + 1) * PAGE_LENGTH))])
            .replace("OUT", "\033[0;33mOUT\033[0m")
            .replace("IN", "\033[0;32mIN\033[0m")
            .replace("NOR", "\033[0;31mNOR\033[0m")
            .replace("@", f"\33[36m@\33[0m")
            .replace("|", f"\33[36m|\33[0m")
            .replace("-", f"\33[36m-\33[0m")
            .replace(datetime.now().strftime("%m/%d/%Y"), f"{stylizedDate}")
        )

    # List out_students (with names)
    if (len(out_students) > 0):
        student_names = []
        for student in out_students:
            student_names.append(f"{student['first']} {student['last']}")
        print(f"\n\033[0;33mCurrently out\033[0m: {', '.join(student_names)}")

    print(f"Total: {len(text)}")
    print("-----\n")

    print("\033[0;33mOUT\033[0m: Left the room")
    print("\033[0;32mIN\033[0m: Returned to the room")
    print(
        "\033[0;31mNOR\033[0m: Forgot to sign back in (or never returned to the room)\n")

    print("-----\n")

    print(f"(\33[36m{ADMIN_CODE}\33[0m to return; \33[36mEnter\33[0m to next page, \33[36m'.'\33[0m to previous page)")
    new_page = input("\nPage: ")

    # Return to selector page
    if new_page.strip() == ADMIN_CODE:
        global curr_mode
        curr_mode = "modes"
        page = 0
        return 0

    # Next page
    if new_page.strip() == "":
        page = min(page + 1, MAX_PAGES - 1)
        return 0

    # Previous page
    if new_page.strip()[0] == ".":
        page = max(page - 1, 0)
        return 0

    # Calculate page based on input
    if (new_page.isdigit() and int(new_page) > 0):
        page = max(min(int(new_page) - 1, MAX_PAGES - 1), 0)
        return 0

    return 0


def mode_csv_view():
    # Current terminal height - 9; sizing that doesnt include surrounding text
    PAGE_LENGTH = os.get_terminal_size().lines - 9

    global page

    # Join user data into a string
    def join(user, color):
        return f"\033[0;{color}m{str(user['user'])[1:]} | {user['last']}, {user['first']}\033[0m\n"

    text = []
    for i, user in enumerate(users["last"]):
        # alternate colors to make it easier to read
        text.append(join(user, "36" if i % 2 == 1 else "0"))

    # Pagination
    MAX_PAGES = math.ceil(len(text)/PAGE_LENGTH)
    print(
        f"CSV Users ({page + 1}/{MAX_PAGES}):\n")
    print(
        "".join(text[slice((page * PAGE_LENGTH), ((page + 1) * PAGE_LENGTH))]))

    print(f"Total: {len(text)}")
    print("-----\n")

    print(f"(\33[36m{ADMIN_CODE}\33[0m to return; \33[36mEnter\33[0m to next page, \33[36m'.'\33[0m to previous page)")
    new_page = input("\nPage: ")

    # Return to selector page
    if new_page.strip() == ADMIN_CODE:
        global curr_mode
        curr_mode = "modes"
        page = 0
        return 0

    # Next page
    if new_page.strip() == "":
        page = min(page + 1, MAX_PAGES - 1)
        return 0

    # Previous page
    if new_page.strip()[0] == ".":
        page = max(page - 1, 0)
        return 0

    # Calculate page based on input
    if (new_page.isdigit() and int(new_page) > 0):
        page = max(min(int(new_page) - 1, MAX_PAGES - 1), 0)
        return 0

    return 0


def mode_modes():
    # Options
    opts = ("Normal (001)", "Admin Logs (002)", "View CSV buffer (003)",
            "Reload CSV (004)", "Reload Out (005)")
    index = survey.routines.select('Modes: ', options=opts)

    # Mode switcher
    global curr_mode
    global page

    # Use index to calculate option
    match "00" + str(index + 1):
        case "001":
            curr_mode = "normal"
        case "002":
            curr_mode = "admin"
            page = 1
        case "003":
            curr_mode = "csv_view"
            page = 0
        case "004":
            global users
            users = readCSV(USER_CSV)
            print("\033[0;32mCSV reloaded\033[0m")

            print(f"\n{len(users['user'])} students loaded")
            return 2
        case "005":
            loadOutStudents()
            print("\033[0;32mOut students reloaded\033[0m")

            # Print currently out students once reloaded
            if (len(out_students) > 0):
                student_names = []
                for student in out_students:
                    student_names.append(
                        student["user"] + " | " + student["first"] + " " + student["last"])
                print(f"\n{', '.join(student_names)}")
                return 5
            else:
                print("\nNone")
                return 2

    return 0


# Modes dictionary
modes = {
    "modes": mode_modes,   # ADMIN_CODE
    "normal": mode_normal,  # 001
    "admin": mode_admin,   # 002
    "csv_view": mode_csv_view,   # 003
}

try:
    os.system("cls")
    while True:
        # Run current mode function defined in the modes dictionary
        status = modes[curr_mode]()

        # Depending on what value is returned from the mode function, it determines what delay is used
        if status == 0:
            os.system('cls')
        elif status:
            time.sleep(status)
            os.system('cls')

        time.sleep(0.01)  # 10ms delay so program doesnt break computer
except KeyboardInterrupt:
    print("\n\033[0;31mProgram was stopped by user\033[0m")
