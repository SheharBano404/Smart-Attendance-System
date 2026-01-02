import cv2
import os
import numpy as np
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import csv
import json
from PIL import Image, ImageTk
import win32com.client
import pythoncom
import time
import threading
import speech_recognition as sr
from queue import Queue, Empty
import re

# ---------------- CONFIG ----------------
FACES_DIR = "faces"
ADMIN_FOLDER = os.path.join(FACES_DIR, "Admin")
TRAINER_FILE = "trainer.yml"
LABEL_MAP_FILE = "label_map.json"
ID_MAP_FILE = "id_map.json"
ATTENDANCE_FILE = "attendance.csv"
BLOCKED_FILE = "blocked.json"
ADMIN_USERNAME = "sheharbano404"

FACE_SIZE = (200, 200)
UNKNOWN_THRESHOLD = 70.0
MIN_FACE_SIZE = (60, 60)
SAMPLES_NEW_USER = 5

# Face-switch / stale-name fix (UPGRADE)
FACE_CHANGE_FRAMES = 4
FACE_CHANGE_DIST_THRESHOLD = 0.55  # Bhattacharyya distance > this => likely different face

os.makedirs(FACES_DIR, exist_ok=True)
os.makedirs(ADMIN_FOLDER, exist_ok=True)

admin_root_ref = None
VOICE_COMMAND_QUEUE = Queue()
SPEECH_QUEUE = Queue()

# ---------------- SPEECH ENGINE ----------------
def speech_worker():
    pythoncom.CoInitialize()
    try:
        engine = win32com.client.Dispatch("SAPI.SpVoice")
        while True:
            text = SPEECH_QUEUE.get()
            if text is None:
                break
            try:
                print(f"Jerry says: {text}")
                engine.Speak(text)
            except Exception as e:
                print(f"Speech error: {e}")
            finally:
                SPEECH_QUEUE.task_done()
    finally:
        pythoncom.CoUninitialize()


speech_thread = threading.Thread(target=speech_worker, daemon=True)
speech_thread.start()


def speak_async(text: str):
    SPEECH_QUEUE.put(text)


def clear_speech_queue():
    """Clear any pending speech to prevent old announcements"""
    while not SPEECH_QUEUE.empty():
        try:
            SPEECH_QUEUE.get_nowait()
            SPEECH_QUEUE.task_done()
        except Empty:
            break


def speak_sync(text: str):
    try:
        print(f"Speaking (sync): {text}")
        temp_voice = win32com.client.Dispatch("SAPI.SpVoice")
        temp_voice.Speak(text)
    except Exception as e:
        print(f"Error in sync speech: {e}")


# ---------------- SPECIAL TEACHER WELCOME ----------------
TEACHER_NAME_ALIASES_RAW = [
    "Dr. Hamedoon",
    "Dr. Syed M Hamedoon",
    "Hamedoon",
    "Syed M Hamedoon",
    "Syed Hamedoon",
    "M Hamedoon",
]


def normalize_person_name(name: str) -> str:
    if name is None:
        return ""
    s = str(name).strip().lower()
    s = re.sub(r"[^\w\s]", " ", s)  # remove punctuation
    s = re.sub(r"\s+", " ", s).strip()
    return s


TEACHER_NAME_ALIASES = {normalize_person_name(n) for n in TEACHER_NAME_ALIASES_RAW}


def is_special_teacher_name(name: str) -> bool:
    return normalize_person_name(name) in TEACHER_NAME_ALIASES


def teacher_welcome_lines(name: str):
    return [
        f"Honorable Sir {name}, your presence shines like a lantern guiding us through the night of ignorance.", 
        "Your wisdom flows like a river, nourishing minds and shaping futures with quiet strength.",
    ]

# ---------------- FACE SIGNATURE (UPGRADE: fixes stale previous user name) ----------------
def compute_face_hist(face_gray_resized: np.ndarray) -> np.ndarray:
    """
    Create a simple robust signature of the face using histogram.
    Used only for detecting "person changed" so we can reset state immediately.
    """
    hist = cv2.calcHist([face_gray_resized], [0], None, [64], [0, 256])
    cv2.normalize(hist, hist)
    return hist


def hist_distance(h1: np.ndarray | None, h2: np.ndarray | None) -> float:
    if h1 is None or h2 is None:
        return 0.0
    try:
        return float(cv2.compareHist(h1, h2, cv2.HISTCMP_BHATTACHARYYA))
    except Exception:
        return 0.0


# ---------------- UTILITIES ----------------
def ensure_attendance_file():
    desired_header = ["Name", "ID", "Date", "Time", "Day", "Blocked"]

    if not os.path.exists(ATTENDANCE_FILE):
        with open(ATTENDANCE_FILE, "w", newline="") as f:
            csv.writer(f).writerow(desired_header)
        return

    try:
        with open(ATTENDANCE_FILE, "r", newline="") as f:
            rows = list(csv.reader(f))
    except Exception:
        rows = []

    if not rows:
        with open(ATTENDANCE_FILE, "w", newline="") as f:
            csv.writer(f).writerow(desired_header)
        return

    header = rows[0]
    if len(header) >= 6 and header[:6] == desired_header:
        return

    new_rows = [desired_header]
    for r in rows[1:]:
        if not r:
            continue
        if r[0] == "Name":
            continue

        base = (r + ["", "", "", "", ""])[:5]
        blocked_val = ""

        if len(r) >= 6:
            sixth = str(r[5]).strip()
            if sixth.upper() == "BLOCKED":
                blocked_val = "BLOCKED"

        new_rows.append(base + [blocked_val])

    with open(ATTENDANCE_FILE, "w", newline="") as f:
        csv.writer(f).writerows(new_rows)


def has_marked_today(name: str) -> bool:
    ensure_attendance_file()
    today = datetime.now().strftime("%Y-%m-%d")
    with open(ATTENDANCE_FILE, "r", newline="") as f:
        for row in csv.reader(f):
            if len(row) >= 3 and row[0] == name and row[2] == today:
                if len(row) >= 6 and str(row[5]).strip().upper() == "BLOCKED":
                    continue
                return True
    return False


def has_blocked_attempt_today(name: str) -> bool:
    ensure_attendance_file()
    today = datetime.now().strftime("%Y-%m-%d")
    with open(ATTENDANCE_FILE, "r", newline="") as f:
        for row in csv.reader(f):
            if (
                len(row) >= 6
                and row[0] == name
                and row[2] == today
                and str(row[5]).strip().upper() == "BLOCKED"
            ):
                return True
    return False


def log_blocked_attempt(name: str, numeric_id: str) -> bool:
    ensure_attendance_file()

    if has_blocked_attempt_today(name):
        return False

    now = datetime.now()
    date = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%I:%M:%S %p")
    day = now.strftime("%A")

    with open(ATTENDANCE_FILE, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([name, numeric_id, date, time_str, day, "BLOCKED"])

    return True


def mark_attendance(name: str, numeric_id: str, is_admin: bool = False, speak: bool = True) -> bool:
    ensure_attendance_file()
    already_marked = has_marked_today(name)

    if not already_marked:
        now = datetime.now()
        date = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%I:%M:%S %p")
        day = now.strftime("%A")
        with open(ATTENDANCE_FILE, "a", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([name, numeric_id, date, time_str, day, ""])

    if speak:
        if not already_marked:
            if is_admin:
                speak_async(f"{name}, mighty admin! Your attendance scroll has been stamped!")
            else:
                speak_async(f"{name}, zap! Your presence has been laser-etched into the system!")
        else:
            speak_async(f"{name}, your attendance was already marked today.")

    return not already_marked


def load_blocked_users():
    if not os.path.exists(BLOCKED_FILE):
        return []
    with open(BLOCKED_FILE, "r") as f:
        return json.load(f)


def save_blocked_users(blocked_list):
    with open(BLOCKED_FILE, "w") as f:
        json.dump(blocked_list, f)


def save_label_map(label_map: dict):
    with open(LABEL_MAP_FILE, "w") as f:
        json.dump(label_map, f)


def load_label_map():
    if not os.path.exists(LABEL_MAP_FILE):
        return {}
    with open(LABEL_MAP_FILE, "r") as f:
        return json.load(f)


def save_id_map(id_map: dict):
    with open(ID_MAP_FILE, "w") as f:
        json.dump(id_map, f)


def load_id_map():
    if not os.path.exists(ID_MAP_FILE):
        return {}
    with open(ID_MAP_FILE, "r") as f:
        return json.load(f)


# ---------------- ATTENDANCE VIEWER (RUNTIME RED/GREEN) -----------------
def view_attendance_gui():
    ensure_attendance_file()

    viewer = tk.Toplevel()
    viewer.title("Sacred Scroll of Attendance")
    viewer.attributes("-topmost", True)

    screen_width = viewer.winfo_screenwidth()
    screen_height = viewer.winfo_screenheight()
    w, h = 1000, 600
    x = (screen_width // 2) - (w // 2)
    y = (screen_height // 2) - (h // 2)
    viewer.geometry(f"{w}x{h}+{x}+{y}")
    viewer.configure(bg="#f0e6f6")

    tk.Label(
        viewer,
        text="📜 The Sacred Attendance Scroll 📜",
        font=("Comic Sans MS", 18, "bold"),
        bg="#f0e6f6",
        fg="#4b0082",
    ).pack(pady=10)

    table_frame = tk.Frame(viewer)
    table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    scroll_y = tk.Scrollbar(table_frame)
    scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

    columns = ("Name", "ID", "Date", "Time", "Day", "Blocked")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings", yscrollcommand=scroll_y.set)

    for col in columns:
        tree.heading(col, text=col)
        if col == "Blocked":
            tree.column(col, width=120, anchor=tk.CENTER)
        else:
            tree.column(col, width=140, anchor=tk.CENTER)

    scroll_y.config(command=tree.yview)
    tree.pack(fill=tk.BOTH, expand=True)

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Comic Sans MS", 12, "bold"))
    style.configure("Treeview", font=("Arial", 11), rowheight=25)

    tree.tag_configure("blocked", background="#ffcccc", foreground="#cc0000")
    tree.tag_configure("present", background="#ccffcc", foreground="#006600")

    row_key_to_file_blocked = {}
    last_mtime = {"v": None}

    def load_rows_if_file_changed():
        ensure_attendance_file()
        try:
            mtime = os.path.getmtime(ATTENDANCE_FILE)
        except Exception:
            mtime = None

        if mtime == last_mtime["v"]:
            return

        last_mtime["v"] = mtime

        try:
            y0, _y1 = tree.yview()
        except Exception:
            y0 = 0.0

        tree.delete(*tree.get_children())
        row_key_to_file_blocked.clear()

        try:
            with open(ATTENDANCE_FILE, "r", newline="") as f:
                reader = csv.reader(f)
                for row in reader:
                    if not row or row[0] == "Name":
                        continue

                    while len(row) < 6:
                        row.append("")

                    name, _id, date, t, day, blocked_cell = row[:6]
                    blocked_cell_norm = "BLOCKED" if str(blocked_cell).strip().upper() == "BLOCKED" else ""
                    row[5] = blocked_cell_norm

                    key = (name, _id, date, t, day)
                    row_key_to_file_blocked[key] = blocked_cell_norm

                    tree.insert("", tk.END, values=row)
        except Exception as e:
            print("Error reading CSV:", e)

        try:
            tree.yview_moveto(y0)
        except Exception:
            pass

    def refresh_runtime_colors():
        load_rows_if_file_changed()

        try:
            blocked_now = set(load_blocked_users())

            for item in tree.get_children():
                vals = list(tree.item(item, "values"))
                if len(vals) < 6:
                    continue

                name, _id, date, t, day = vals[0], vals[1], vals[2], vals[3], vals[4]
                file_blocked = row_key_to_file_blocked.get((name, _id, date, t, day), "")
                file_blocked = "BLOCKED" if str(file_blocked).strip().upper() == "BLOCKED" else ""

                if name in blocked_now:
                    vals[5] = "BLOCKED"
                    tree.item(item, values=vals, tags=("blocked",))
                else:
                    vals[5] = file_blocked
                    if vals[5] == "BLOCKED":
                        tree.item(item, values=vals, tags=("blocked",))
                    else:
                        tree.item(item, values=vals, tags=("present",))

        except Exception as e:
            print("Runtime refresh error:", e)

        if viewer.winfo_exists():
            viewer.after(1000, refresh_runtime_colors)

    load_rows_if_file_changed()
    refresh_runtime_colors()

    tk.Button(
        viewer,
        text="Close Scroll",
        command=viewer.destroy,
        bg="#ff7f50",
        fg="#4b0082",
        font=("Comic Sans MS", 12, "bold"),
        relief="raised",
        bd=4,
    ).pack(pady=15)


# ---------------- STYLISH DIALOGS ----------------
def stylish_askstring(title, prompt, show=""):
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.attributes("-topmost", True)
    dialog.configure(bg="#f0e6f6")
    dialog.geometry("420x220")
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width // 2) - (420 // 2)
    y = (screen_height // 2) - (220 // 2)
    dialog.geometry(f"420x220+{x}+{y}")
    dialog.resizable(False, False)
    font = ("Comic Sans MS", 12, "bold")

    tk.Label(dialog, text=prompt, bg="#f0e6f6", fg="#4b0082", font=font).pack(pady=15)
    entry = tk.Entry(
        dialog,
        show=show,
        font=font,
        bg="#fff0f5",
        fg="#4b0082",
        insertbackground="#4b0082",
        relief="groove",
        bd=3,
    )
    entry.pack(pady=5, ipadx=10, ipady=5)
    entry.focus()
    entry.focus_force()

    result = [None]
    btn_frame = tk.Frame(dialog, bg="#f0e6f6")
    btn_frame.pack(pady=15)

    tk.Button(
        btn_frame,
        text="Submit",
        command=lambda: [result.__setitem__(0, entry.get()), dialog.destroy()],
        bg="#ffb6c1",
        fg="#4b0082",
        font=font,
        relief="raised",
        bd=4,
        width=10,
    ).pack(side=tk.LEFT, padx=15)

    tk.Button(
        btn_frame,
        text="Cancel",
        command=lambda: [result.__setitem__(0, None), dialog.destroy()],
        bg="#87cefa",
        fg="#4b0082",
        font=font,
        relief="raised",
        bd=4,
        width=10,
    ).pack(side=tk.RIGHT, padx=15)

    dialog.wait_window()
    return result[0]


def stylish_messagebox(title, message, type="info"):
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.attributes("-topmost", True)
    dialog.geometry("420x180")
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width // 2) - (420 // 2)
    y = (screen_height // 2) - (180 // 2)
    dialog.geometry(f"420x180+{x}+{y}")
    dialog.resizable(False, False)
    dialog.configure(bg="#e0f7fa")
    font = ("Comic Sans MS", 12, "bold")

    tk.Label(dialog, text=message, bg="#e0f7fa", fg="#00796b", font=font, wraplength=380).pack(pady=20)
    btn_frame = tk.Frame(dialog, bg="#e0f7fa")
    btn_frame.pack(pady=10)

    if type == "yesno":
        result = [None]
        tk.Button(
            btn_frame,
            text="Yes",
            command=lambda: [result.__setitem__(0, True), dialog.destroy()],
            bg="#ffcc80",
            fg="#4b0082",
            font=font,
            relief="raised",
            bd=4,
            width=10,
        ).pack(side=tk.LEFT, padx=10)
        tk.Button(
            btn_frame,
            text="No",
            command=lambda: [result.__setitem__(0, False), dialog.destroy()],
            bg="#90ee90",
            fg="#4b0082",
            font=font,
            relief="raised",
            bd=4,
            width=10,
        ).pack(side=tk.RIGHT, padx=10)
        dialog.wait_window()
        return result[0]
    else:
        tk.Button(
            btn_frame,
            text="OK",
            command=dialog.destroy,
            bg="#ffb6c1",
            fg="#4b0082",
            font=font,
            relief="raised",
            bd=4,
            width=10,
        ).pack()
        dialog.wait_window()


# ---------------- TRAINING ----------------
def build_dataset_and_label_map():
    faces, labels = [], []
    name_to_label, label_map = {}, {}
    next_label = 0

    admin_images = [f for f in os.listdir(ADMIN_FOLDER) if f.lower().endswith((".jpg", ".png", ".jpeg"))]
    if admin_images:
        name_to_label[ADMIN_USERNAME] = 0
        label_map["0"] = ADMIN_USERNAME
        next_label = 1
        for fname in admin_images:
            path = os.path.join(ADMIN_FOLDER, fname)
            img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
            if img is None:
                continue
            try:
                img_resized = cv2.resize(img, FACE_SIZE)
            except Exception:
                continue
            faces.append(img_resized)
            labels.append(0)

    for name in sorted(os.listdir(FACES_DIR)):
        if name == "Admin":
            continue
        person_dir = os.path.join(FACES_DIR, name)
        if not os.path.isdir(person_dir):
            continue
        if name not in name_to_label:
            name_to_label[name] = next_label
            label_map[str(next_label)] = name
            next_label += 1
        label = name_to_label[name]
        for fname in sorted(os.listdir(person_dir)):
            if not fname.lower().endswith((".jpg", ".png", ".jpeg")):
                continue
            path = os.path.join(person_dir, fname)
            img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
            if img is None:
                continue
            try:
                img_resized = cv2.resize(img, FACE_SIZE)
            except Exception:
                continue
            faces.append(img_resized)
            labels.append(label)

    return faces, labels, label_map


def train_and_save_recognizer():
    faces, labels, label_map = build_dataset_and_label_map()
    if len(faces) == 0:
        return None, {}
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    recognizer.train(faces, np.array(labels))
    recognizer.save(TRAINER_FILE)
    save_label_map(label_map)
    speak_async(f"Brain upgrade complete! I munched {len(faces)} face-cookies!")
    return recognizer, label_map


def load_recognizer_and_map():
    label_map = load_label_map()
    if not os.path.exists(TRAINER_FILE) or len(label_map) == 0:
        return None, {}
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    recognizer.read(TRAINER_FILE)
    return recognizer, label_map


# ---------------- REGISTRATION ----------------
def register_user(face_images):
    id_map = load_id_map()
    speak_async("Who dares enter? Reveal thy name, mysterious stranger!")
    name = stylish_askstring("Register", "Enter full name:")
    speak_async("Type your secret numeric code, brave soul!")
    numeric_id = stylish_askstring("Register", "Enter numeric ID (numbers only):")

    if name and numeric_id and numeric_id.isnumeric():
        person_dir = os.path.join(FACES_DIR, name)
        os.makedirs(person_dir, exist_ok=True)
        ts = int(datetime.now().timestamp())
        for i, img in enumerate(face_images):
            cv2.imwrite(os.path.join(person_dir, f"{ts}_{i}.jpg"), img)

        id_map[name] = numeric_id
        save_id_map(id_map)
        train_and_save_recognizer()

        mark_attendance(name, numeric_id, is_admin=False, speak=False)

        # Prevent old queued speech from playing after registration
        clear_speech_queue()

        if is_special_teacher_name(name):
            stylish_messagebox(
                "Welcome Sir",
                f"Welcome Sir {name}!\n\nYour registration is completed."
            )
            for line in teacher_welcome_lines(name):
                speak_async(line)
            speak_async("Your registration and attendance have been marked successfully.")
        else:
            speak_async(f"{name}, you have been registered and your attendance scroll has been stamped!")

        return name, numeric_id

    stylish_messagebox("Error", "Invalid name or numeric ID")
    speak_async("Oops! Registration spell failed")
    return None, None


# ---------------- BLOCK / UNBLOCK ----------------
def block_user():
    id_map = load_id_map()
    blocked = load_blocked_users()
    names = [n for n in id_map.keys() if n not in blocked]
    if not names:
        stylish_messagebox("Info", "No users to block")
        speak_async("No mortals left to banish!")
        return

    speak_async("Casting a digital ban-spell... who shall be cursed?")
    user_to_block = stylish_askstring("Block User", f"Enter name to block:\nAvailable: {', '.join(names)}")
    if user_to_block and user_to_block in names:
        blocked.append(user_to_block)
        save_blocked_users(blocked)
        stylish_messagebox("Blocked", f"User {user_to_block} blocked")
        speak_async(f"User {user_to_block} has been banished to the shadow realm!")


def unblock_user():
    blocked = load_blocked_users()
    if not blocked:
        stylish_messagebox("Info", "No users are blocked")
        speak_async("No souls are cursed right now!")
        return

    speak_async("Lifting the curse... who shall be freed?")
    user_to_unblock = stylish_askstring("Unblock User", f"Enter name to unblock:\nBlocked: {', '.join(blocked)}")
    if user_to_unblock and user_to_unblock in blocked:
        blocked.remove(user_to_unblock)
        save_blocked_users(blocked)
        stylish_messagebox("Unblocked", f"User {user_to_unblock} unblocked")
        speak_async(f"User {user_to_unblock} has been freed from the digital dungeon!")


# ---------------- CAMERA & ATTENDANCE (UPGRADED: fixes "previous name repeats") ----------------
def camera_for_one_user(is_admin=False):
    recognizer, label_map = load_recognizer_and_map()
    label_map_int = {int(k): v for k, v in label_map.items()} if label_map else {}
    id_map = load_id_map()

    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        stylish_messagebox("Error", "Cannot open camera")
        speak_async("Oops! The magic eye refuses to wake up")
        return

    # Clear any stale voice output to avoid caching/old prompts
    clear_speech_queue()
    speak_async("Opening the crystal ball of attendance...")
    speak_async("Camera is open. Press Q to close.")

    unknown_faces_buffer = []
    buffer_count = 0

    # ===== STATE TRACKING =====
    current_confirmed_name = None
    last_predicted_label = None
    recognition_stable_count = 0
    STABLE_THRESHOLD = 5
    announcement_made = False
    attendance_marked_current = False

    no_face_frames = 0
    NO_FACE_RESET_THRESHOLD = 10

    unknown_stable_count = 0
    blocked_logged_current = False

    # ===== NEW: FACE SWITCH DETECTION =====
    prev_hist = None
    switch_count = 0
    # =====================================

    camera_window = tk.Toplevel()
    camera_window.attributes("-topmost", True)
    camera_window.title("Camera Feed")

    screen_width = camera_window.winfo_screenwidth()
    screen_height = camera_window.winfo_screenheight()
    w, h = 640, 480
    x = (screen_width // 2) - (w // 2)
    y = (screen_height // 2) - (h // 2)
    camera_window.geometry(f"{w}x{h}+{x}+{y}")

    label = tk.Label(camera_window)
    label.pack()

    camera_window.lift()
    camera_window.focus_force()

    stop_camera = [False]

    def reset_subject_state():
        nonlocal current_confirmed_name, last_predicted_label, recognition_stable_count
        nonlocal announcement_made, attendance_marked_current, unknown_stable_count
        nonlocal unknown_faces_buffer, buffer_count, blocked_logged_current
        nonlocal prev_hist, switch_count

        current_confirmed_name = None
        last_predicted_label = None
        recognition_stable_count = 0
        announcement_made = False
        attendance_marked_current = False
        unknown_stable_count = 0
        unknown_faces_buffer.clear()
        buffer_count = 0
        blocked_logged_current = False

        # Important: stop old queued lines from previous person
        clear_speech_queue()

        # keep prev_hist as-is to avoid instant re-trigger loop
        switch_count = 0

    def on_close():
        stop_camera[0] = True
        camera_window.destroy()

    camera_window.protocol("WM_DELETE_WINDOW", on_close)
    camera_window.bind("<KeyPress-q>", lambda e: on_close())
    camera_window.bind("<KeyPress-Q>", lambda e: on_close())

    try:
        while not stop_camera[0]:
            ret, frame = cap.read()
            if not ret:
                break

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, 1.1, 5, minSize=MIN_FACE_SIZE)
            display = frame.copy()

            if len(faces) == 0:
                no_face_frames += 1
                if no_face_frames >= NO_FACE_RESET_THRESHOLD:
                    reset_subject_state()
                    prev_hist = None
                rgb_display = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(rgb_display)
                imgtk = ImageTk.PhotoImage(image=img)
                label.config(image=imgtk)
                label.imgtk = imgtk
                camera_window.update_idletasks()
                camera_window.update()
                time.sleep(0.01)
                continue
            else:
                no_face_frames = 0

            # ---- IMPORTANT UPGRADE:
            # Process ONLY the biggest face to avoid mixing states with multiple faces.
            fx, fy, fw, fh = max(faces, key=lambda r: r[2] * r[3])

            face_gray = gray[fy:fy + fh, fx:fx + fw]
            try:
                face_small = cv2.resize(face_gray, FACE_SIZE)
            except Exception:
                continue

            # Face switch detection (new person came without removing face completely)
            cur_hist = compute_face_hist(face_small)
            if prev_hist is not None:
                dist = hist_distance(cur_hist, prev_hist)
                if dist > FACE_CHANGE_DIST_THRESHOLD:
                    switch_count += 1
                else:
                    switch_count = 0

                if switch_count >= FACE_CHANGE_FRAMES:
                    # New person likely -> reset immediately so old name never speaks
                    reset_subject_state()
            prev_hist = cur_hist

            name_text = "Detecting..."
            color = (255, 255, 0)
            is_known = False
            is_admin_detected = False

            blocked_users = load_blocked_users()  # refresh live

            if recognizer is not None and len(label_map_int) > 0:
                try:
                    pred_label, conf = recognizer.predict(face_small)

                    # If predicted label changes, treat as a new subject
                    if last_predicted_label is not None and pred_label != last_predicted_label:
                        reset_subject_state()

                    last_predicted_label = pred_label

                    if pred_label in label_map_int and conf < UNKNOWN_THRESHOLD:
                        detected_name = label_map_int[pred_label]
                        unknown_stable_count = 0

                        if current_confirmed_name == detected_name:
                            recognition_stable_count += 1
                        else:
                            current_confirmed_name = detected_name
                            recognition_stable_count = 1
                            announcement_made = False
                            attendance_marked_current = False
                            blocked_logged_current = False
                            clear_speech_queue()

                        if recognition_stable_count >= STABLE_THRESHOLD:
                            name_text = detected_name
                            color = (0, 255, 127)
                            is_known = True
                            is_admin_detected = (detected_name == ADMIN_USERNAME)
                            numeric_id = id_map.get(detected_name, "UnknownID")

                            if detected_name in blocked_users:
                                color = (0, 0, 255)
                                name_text = f"{detected_name} [BLOCKED]"

                                if not blocked_logged_current:
                                    log_blocked_attempt(detected_name, numeric_id)
                                    blocked_logged_current = True

                                if not announcement_made:
                                    speak_async(f"{detected_name} has been cursed and cannot enter! Access denied!")
                                    announcement_made = True
                            else:
                                if not attendance_marked_current:
                                    just_marked_now = mark_attendance(
                                        ADMIN_USERNAME if is_admin_detected else detected_name,
                                        "ADMINID" if is_admin_detected else numeric_id,
                                        is_admin=is_admin_detected,
                                        speak=False,
                                    )
                                    attendance_marked_current = True

                                    if not announcement_made:
                                        if just_marked_now:
                                            if is_admin_detected:
                                                speak_async(f"{detected_name}, mighty admin! Your attendance scroll has been stamped!")
                                            else:
                                                speak_async(f"{detected_name}, zap! Your presence has been laser-etched into the system!")
                                        else:
                                            # NOTE: we only say "already marked" for the SAME person.
                                            if is_admin_detected:
                                                speak_async(f"{detected_name}, mighty admin, your attendance was already marked today.")
                                            else:
                                                speak_async(f"{detected_name}, your attendance was already marked today.")
                                        announcement_made = True
                        else:
                            name_text = "Verifying..."
                            color = (255, 255, 0)

                    else:
                        # Unknown or low confidence
                        if current_confirmed_name is not None:
                            # drop old confirmed identity immediately
                            reset_subject_state()

                        unknown_stable_count += 1
                        name_text = "Unknown"
                        color = (255, 182, 193)

                except Exception as e:
                    print(f"Recognition error: {e}")
            else:
                # No recognizer yet -> treat as unknown
                unknown_stable_count += 1
                name_text = "Unknown"
                color = (255, 182, 193)

            # Unknown registration (works now even when users switch quickly)
            if (not is_known) and (not is_admin):
                if unknown_stable_count >= STABLE_THRESHOLD:
                    if buffer_count < SAMPLES_NEW_USER:
                        unknown_faces_buffer.append(face_small)
                        buffer_count += 1
                        cv2.putText(
                            display,
                            "Collecting magical face data...",
                            (fx, max(20, fy - 30)),
                            cv2.FONT_HERSHEY_SIMPLEX,
                            0.7,
                            (173, 216, 230),
                            2,
                        )
                    if buffer_count >= SAMPLES_NEW_USER:
                        _name, _numeric_id = register_user(unknown_faces_buffer)
                        unknown_faces_buffer.clear()
                        buffer_count = 0
                        cap.release()
                        if camera_window.winfo_exists():
                            camera_window.destroy()
                        return

            # Draw only one face box (largest)
            cv2.rectangle(display, (fx, fy), (fx + fw, fy + fh), color, 2)
            cv2.putText(display, name_text, (fx, fy - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)

            rgb_display = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb_display)
            imgtk = ImageTk.PhotoImage(image=img)
            label.config(image=imgtk)
            label.imgtk = imgtk

            camera_window.lift()
            camera_window.focus_force()

            camera_window.update_idletasks()
            camera_window.update()

            time.sleep(0.01)

    finally:
        cap.release()
        if camera_window.winfo_exists():
            camera_window.destroy()


# ---------------- ADMIN FACE CAPTURE (TKINTER VERSION) ----------------
def register_admin_face():
    admin_img_path = os.path.join(ADMIN_FOLDER, "admin.jpg")
    if os.path.exists(admin_img_path):
        speak_async("Admin face already registered in the vault of faces!")
        return True

    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        stylish_messagebox("Error", "Cannot open camera")
        speak_async("Oops! The magic eye refuses to wake up")
        return False

    clear_speech_queue()
    speak_async("Look at the magic eye. Press S to capture your legendary admin mugshot!")
    speak_async("Press Q to close.")

    camera_window = tk.Toplevel()
    camera_window.attributes("-topmost", True)
    camera_window.title("Admin Face Capture - Press S")

    screen_width = camera_window.winfo_screenwidth()
    screen_height = camera_window.winfo_screenheight()
    w, h = 640, 480
    x = (screen_width // 2) - (w // 2)
    y = (screen_height // 2) - (h // 2)
    camera_window.geometry(f"{w}x{h}+{x}+{y}")

    label = tk.Label(camera_window)
    label.pack()

    camera_window.lift()
    camera_window.focus_force()

    stop_camera = [False]
    captured = [False]
    faces = []

    def on_close():
        stop_camera[0] = True
        camera_window.destroy()

    camera_window.protocol("WM_DELETE_WINDOW", on_close)
    camera_window.bind("<KeyPress-q>", lambda e: on_close())
    camera_window.bind("<KeyPress-Q>", lambda e: on_close())

    def capture_face(event):
        if not captured[0] and len(faces) > 0:
            ret2, frame2 = cap.read()
            if ret2:
                gray2 = cv2.cvtColor(frame2, cv2.COLOR_BGR2GRAY)
                x0, y0, w0, h0 = faces[0]
                face_img = gray2[y0:y0 + h0, x0:x0 + w0]
                face_resized = cv2.resize(face_img, FACE_SIZE)
                cv2.imwrite(admin_img_path, face_resized)
                speak_async("Gotcha! Admin mugshot stored in the digital vault!")
                stylish_messagebox("Success", "Admin face saved")
                captured[0] = True
                stop_camera[0] = True

    camera_window.bind("<KeyPress-s>", capture_face)
    camera_window.bind("<KeyPress-S>", capture_face)

    try:
        while not stop_camera[0]:
            ret, frame = cap.read()
            if not ret:
                break

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, 1.1, 5, minSize=(60, 60))
            display = frame.copy()

            for (x1, y1, w1, h1) in faces:
                cv2.rectangle(display, (x1, y1), (x1 + w1, y1 + h1), (0, 255, 0), 2)

            rgb_display = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb_display)
            imgtk = ImageTk.PhotoImage(image=img)

            label.config(image=imgtk)
            label.imgtk = imgtk

            camera_window.lift()
            camera_window.focus_force()

            camera_window.update_idletasks()
            camera_window.update()

            time.sleep(0.01)

    finally:
        cap.release()
        if camera_window.winfo_exists():
            camera_window.destroy()
        cv2.destroyAllWindows()

    if captured[0]:
        train_and_save_recognizer()
        return True

    speak_async("Admin face capture cancelled... the vault remains empty!")
    return False


# ---------------- ADMIN PANEL CLOSE ----------------
def close_admin_panel():
    global admin_root_ref
    if admin_root_ref is not None and admin_root_ref.winfo_exists():
        speak_async("Sealing the vault... admin powers fading...")
        admin_root_ref.destroy()
    else:
        speak_async("The admin vault is already sealed.")
    admin_root_ref = None


# ---------------- ADMIN PANEL ----------------
def admin_panel():
    global admin_root_ref

    if admin_root_ref is not None and admin_root_ref.winfo_exists():
        speak_async("Admin vault is already open.")
        admin_root_ref.lift()
        return

    speak_async("Admin login spell activated! Whisper your username...")
    username = stylish_askstring("Admin Login", "Enter Admin Username:")
    speak_async("Now, chant your secret passkey...")
    password = stylish_askstring("Admin Login", "Enter Passkey:", show="*")

    if username != ADMIN_USERNAME or password != "snape":
        stylish_messagebox("Access Denied", "Invalid credentials!")
        speak_async("Access denied! The vault rejects impostors!")
        return

    speak_async("Hail to the code-chief! Admin powers unlocked!")
    register_admin_face()

    admin_root = tk.Toplevel()
    admin_root_ref = admin_root
    admin_root.title("Admin Panel - Vault")
    admin_root.attributes("-fullscreen", True)
    admin_root.attributes("-topmost", True)
    admin_root.configure(bg="#fff5ee")
    admin_root.bind("<Escape>", lambda e: close_admin_panel())

    font = ("Comic Sans MS", 20, "bold")
    admin_frame = tk.Frame(admin_root, bg="#fff5ee")
    admin_frame.place(relx=0.5, rely=0.5, anchor="center")

    tk.Label(
        admin_frame,
        text="Ye Admin Vault",
        bg="#fff5ee",
        fg="#4b0082",
        font=("Comic Sans MS", 30, "bold"),
    ).pack(pady=30)

    def create_btn(text, cmd, bgc, fgc):
        return tk.Button(
            admin_frame,
            text=text,
            command=cmd,
            bg=bgc,
            fg=fgc,
            font=font,
            relief="raised",
            bd=5,
            width=25,
            height=2,
            activebackground="#ffd700",
            activeforeground="#4b0082",
        )

    create_btn(
        "View Attendance",
        lambda: [speak_async("Summoning the sacred scroll of attendance..."), view_attendance_gui()],
        "#ffb347",
        "#4b0082",
    ).pack(pady=10)

    create_btn(
        "Block User",
        lambda: [speak_async("Casting a ban-spell..."), block_user()],
        "#87cefa",
        "#4b0082",
    ).pack(pady=10)

    create_btn(
        "Unblock User",
        lambda: [speak_async("Breaking the curse..."), unblock_user()],
        "#90ee90",
        "#4b0082",
    ).pack(pady=10)

    create_btn(
        "Mark Admin Attendance",
        lambda: [speak_async("Admin roll-call rocket launching!"), camera_for_one_user(is_admin=True)],
        "#ffb6c1",
        "#4b0082",
    ).pack(pady=10)

    create_btn("Close Panel", close_admin_panel, "#ff7f50", "#4b0082").pack(pady=15)

    admin_root.protocol("WM_DELETE_WINDOW", close_admin_panel)
    speak_async("Admin vault opened... proceed with caution!")


# ---------------- ATTENDANCE WORKFLOW ----------------
def attendance_workflow(skip_initial_prompt=False):
    if not skip_initial_prompt:
        speak_async("Do you wish to ignite the roll-call rocket?")
        ans = stylish_messagebox(
            "Start Attendance",
            "Open camera to mark attendance?\n(Admins use Admin Panel)",
            type="yesno",
        )
        if not ans:
            speak_async("Abort mission! Roll-call rocket stays grounded.")
            return
        speak_async("Booting up the roll-call rocket boosters!")
    else:
        speak_async("Starting attendance directly on your command!")

    while True:
        camera_for_one_user(is_admin=False)
        speak_async("Is another brave soul ready for roll-call?")
        cont = stylish_messagebox("Next Person", "Does anyone else want to mark attendance?", type="yesno")
        if not cont:
            speak_async("Roll-call rocket has landed. Mission complete!")
            break


# ---------------- PROCESS VOICE COMMANDS ----------------
def process_voice_commands(root):
    global admin_root_ref
    try:
        while True:
            cmd = VOICE_COMMAND_QUEUE.get_nowait()
            print("Processing voice command:", cmd)

            if cmd == "start_attendance":
                attendance_workflow(skip_initial_prompt=True)

            elif cmd == "open_admin_panel":
                admin_panel()

            elif cmd == "exit_system":
                speak_sync("Exiting the system. Code-bye!")
                root.destroy()
                return

            elif cmd == "view_attendance":
                if admin_root_ref is None or not admin_root_ref.winfo_exists():
                    speak_async("Open the admin panel first to view attendance using voice.")
                else:
                    speak_async("Summoning the sacred scroll of attendance...")
                    view_attendance_gui()

            elif cmd == "block_user":
                if admin_root_ref is None or not admin_root_ref.winfo_exists():
                    speak_async("Open the admin panel and authenticate before blocking users.")
                else:
                    speak_async("Blocking a user by your command.")
                    block_user()

            elif cmd == "unblock_user":
                if admin_root_ref is None or not admin_root_ref.winfo_exists():
                    speak_async("Open the admin panel and authenticate before unblocking users.")
                else:
                    speak_async("Unblocking a user by your command.")
                    unblock_user()

            elif cmd == "mark_admin_attendance":
                if admin_root_ref is None or not admin_root_ref.winfo_exists():
                    speak_async("Open the admin panel and authenticate before marking admin attendance.")
                else:
                    speak_async("Marking admin attendance by your command.")
                    camera_for_one_user(is_admin=True)

            elif cmd == "exit_admin_panel":
                close_admin_panel()

    except Empty:
        pass

    try:
        if root.winfo_exists():
            root.after(200, lambda: process_voice_commands(root))
    except Exception:
        pass


# ---------------- VOICE COMMAND LISTENER ----------------
def start_voice_listener(root):
    def listener_loop():
        recognizer = sr.Recognizer()
        try:
            mic = sr.Microphone()
        except Exception as e:
            print("Microphone not available or error initializing microphone:", e)
            return

        with mic as source:
            try:
                print("Calibrating microphone for ambient noise...")
                recognizer.adjust_for_ambient_noise(source, duration=1)
            except Exception as e:
                print("Mic calibration error:", e)

        speak_async(
            "Voice control ready. Say 'Hello' to wake me, "
            "then give your command. Say 'Over' to stop listening."
        )

        listening_mode = False

        while True:
            try:
                with mic as source:
                    print("Listening...")
                    audio = recognizer.listen(source, phrase_time_limit=5)

                try:
                    text = recognizer.recognize_google(audio, language="en-US")
                except sr.UnknownValueError:
                    continue
                except sr.RequestError as e:
                    print("Speech recognition request error:", e)
                    time.sleep(2)
                    continue

                if not text:
                    continue

                text_lower = text.lower().strip()
                print("Heard:", text_lower)

                if not listening_mode:
                    if "hello" not in text_lower:
                        continue
                    listening_mode = True
                    speak_async("Jerry here! I'm listening. Say 'Over' when you're done.")

                if "over" in text_lower:
                    speak_async("Over and out. Going quiet.")
                    listening_mode = False
                    continue

                if ("exit admin" in text_lower or "close admin" in text_lower or "quit admin" in text_lower):
                    VOICE_COMMAND_QUEUE.put("exit_admin_panel")

                elif ("open admin panel" in text_lower or "admin panel" in text_lower):
                    VOICE_COMMAND_QUEUE.put("open_admin_panel")

                elif ("start" in text_lower or "start attendance" in text_lower or "start attendnace" in text_lower):
                    VOICE_COMMAND_QUEUE.put("start_attendance")

                elif ("exit the system" in text_lower or "exit system" in text_lower or "close the system" in text_lower or "quit the system" in text_lower):
                    VOICE_COMMAND_QUEUE.put("exit_system")

                elif "view attendance" in text_lower or "show attendance" in text_lower or "show" in text_lower or "view" in text_lower:
                    VOICE_COMMAND_QUEUE.put("view_attendance")

                elif "block user" in text_lower or "block a user" in text_lower:
                    VOICE_COMMAND_QUEUE.put("block_user")

                elif "unblock user" in text_lower or "unblock" in text_lower or "un block" in text_lower:
                    VOICE_COMMAND_QUEUE.put("unblock_user")

                elif ("mark" in text_lower or "mark admin attendance" in text_lower or "mark my attendance" in text_lower):
                    VOICE_COMMAND_QUEUE.put("mark_admin_attendance")

            except Exception as e:
                print("Voice listener error:", e)
                time.sleep(1)

    t = threading.Thread(target=listener_loop, daemon=True)
    t.start()

    root.after(200, lambda: process_voice_commands(root))


# ---------------- MAIN GUI ----------------
def main_gui():
    ensure_attendance_file()
    root = tk.Tk()
    root.title("Attendance System - Captain Sheharbano's Pirate Ship")
    root.attributes("-fullscreen", True)
    root.attributes("-topmost", True)
    root.configure(bg="#ffe4e1")
    root.bind("<Escape>", lambda event: root.destroy())

    font = ("Comic Sans MS", 20, "bold")
    main_frame = tk.Frame(root, bg="#ffe4e1")
    main_frame.place(relx=0.5, rely=0.5, anchor="center")

    tk.Label(
        main_frame,
        text="Ahoy, Code-Pirates!\nAttendance Ship Ahoy!",
        bg="#ffe4e1",
        fg="#4b0082",
        font=("Comic Sans MS", 30, "bold"),
    ).pack(pady=40)

    def create_main_btn(text, cmd, bgc, fgc):
        return tk.Button(
            main_frame,
            text=text,
            command=cmd,
            bg=bgc,
            fg=fgc,
            font=font,
            relief="raised",
            bd=5,
            width=25,
            height=2,
            activebackground="#ffd700",
            activeforeground="#4b0082",
        )

    create_main_btn(
        "Start Attendance",
        lambda: [speak_async("Booting roll-call rocket boosters!"), attendance_workflow()],
        "#ffb6c1",
        "#4b0082",
    ).pack(pady=15)

    create_main_btn(
        "Admin Panel",
        lambda: [speak_async("Cracking open secret admin vault..."), admin_panel()],
        "#87cefa",
        "#4b0082",
    ).pack(pady=15)

    create_main_btn(
        "Exit",
        lambda: [speak_sync("Shutting down the matrix... Code-bye!"), root.destroy()],
        "#90ee90",
        "#4b0082",
    ).pack(pady=15)

    start_voice_listener(root)
    speak_async("Ahoy, code-pirates! Attendance ship is sailing!")
    root.mainloop()


# ---------------- ENTRY POINT ----------------
if __name__ == "__main__":
    black_dialog = tk.Tk()
    black_dialog.attributes("-fullscreen", True)
    black_dialog.attributes("-topmost", True)
    black_dialog.configure(bg="black")
    black_dialog.overrideredirect(True)

    tk.Label(
        black_dialog,
        text="JERRY",
        bg="black",
        fg="white",
        font=("Comic Sans MS", 60, "bold"),
    ).place(relx=0.5, rely=0.4, anchor="center")

    tk.Label(
        black_dialog,
        text="powered by Shehar's Magic",
        bg="black",
        fg="#ffb6c1",
        font=("Comic Sans MS", 24, "bold"),
    ).place(relx=0.5, rely=0.52, anchor="center")

    black_dialog.update()
    speak_sync("My name is Jerry... summoned into existence by Sherry. Let's wander through my little world together.")
    black_dialog.destroy()
    main_gui()
