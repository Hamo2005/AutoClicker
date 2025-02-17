import pyautogui
import keyboard
import threading
import time
import tkinter as tk
from tkinter import simpledialog, messagebox
from pynput import mouse

# Global variables
recording = False
recorded_events = []

# Record mouse position and clicks
def record_mouse():
    global recording, recorded_events
    recorded_events.clear()

    def on_click(x, y, button, pressed):
        if recording and pressed:
            action = 'Left Click' if button == mouse.Button.left else 'Right Click'
            recorded_events.append((x, y, action, int(time.time() * 1000)))

    listener = mouse.Listener(on_click=on_click)
    listener.start()

    while recording:
        time.sleep(0.01)  # Reduce CPU usage

    listener.stop()

# Save the recorded events
def save_recording():
    if not recorded_events:
        messagebox.showwarning("No Data", "No mouse events recorded.")
        return

    file_name = simpledialog.askstring("Save Recording", "Enter file name:")
    if not file_name:
        return

    try:
        with open(f"{file_name}.txt", 'w') as f:
            start_time = recorded_events[0][3]  # Get the initial timestamp
            for x, y, action, timestamp in recorded_events:
                delay = timestamp - start_time
                f.write(f"{x}, {y}, {action}, {delay}\n")
                start_time = timestamp

        messagebox.showinfo("Success", f"Recording saved to {file_name}.txt")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save: {e}")

# Toggle recording
def toggle_recording():
    global recording

    if recording:
        recording = False
        record_button.config(text="Record")
        save_recording()
    else:
        def countdown():
            for i in range(3, 0, -1):
                record_button.config(text=str(i))
                time.sleep(1)
            record_button.config(text="Recording")
            start_recording()

        threading.Thread(target=countdown, daemon=True).start()

def start_recording():
    global recording
    recording = True
    threading.Thread(target=record_mouse, daemon=True).start()

# GUI setup
def main():
    global recording, record_button

    root = tk.Tk()
    root.title("Mouse Recorder")

    record_button = tk.Button(root, text="Record", command=toggle_recording)
    record_button.pack(padx=20, pady=20)

    def on_close():
        global recording
        recording = False
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    # Stop recording on ESC key
    keyboard.add_hotkey("esc", lambda: toggle_recording() if recording else None)

    root.mainloop()

if __name__ == "__main__":
    main()
