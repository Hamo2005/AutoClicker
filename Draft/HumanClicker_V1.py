import pyautogui
import time
import keyboard
import os
import random
import math
import ctypes
import sys
import traceback
from win32api import GetSystemMetrics

# Configuration
INPUT_FILE = r"E:\AutoClicker\Draft\Clicks_V1.txt"
STOP_KEY = "esc"
CLICK_DELAY = 0.1
GAME_WINDOW_TITLE = "Your Game Window"
HUMAN_MOVE_DURATION = (0.01, 0.1)  # Random movement duration range
BASE_POINTS = 2  # Minimum number of points
MAX_POINTS = 7   # Maximum number of points for long distances
JITTER_AMOUNT = 0.7  # Reduce random jitter to prevent excessive slattering
MICRO_MOVEMENT_CHANCE = 0.03  # Less frequent micro-movements

# Check for admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

def calculate_num_points(start, end):
    """Dynamically adjust the number of points based on distance."""
    distance = math.dist(start, end)
    return min(MAX_POINTS, max(BASE_POINTS, int(distance / 10)))  # Scale based on distance


# Bézier curve implementation
def human_curve(start, end):
    """Generates a smooth Bezier curve path with fewer points for small distances."""
    num_points = calculate_num_points(start, end)

    cp = [
        (start[0] + (end[0] - start[0]) * 0.25, start[1] + (end[1] - start[1]) * 0.25),
        (start[0] + (end[0] - start[0]) * 0.75, start[1] + (end[1] - start[1]) * 0.75)
    ]

    points = []
    for t in [i/num_points for i in range(num_points+1)]:
        x = (1-t)**3 * start[0] + 3*(1-t)**2*t*cp[0][0] + 3*(1-t)*t**2*cp[1][0] + t**3*end[0]
        y = (1-t)**3 * start[1] + 3*(1-t)**2*t*cp[0][1] + 3*(1-t)*t**2*cp[1][1] + t**3*end[1]

        points.append((x + random.uniform(-JITTER_AMOUNT, JITTER_AMOUNT),
                       y + random.uniform(-JITTER_AMOUNT, JITTER_AMOUNT)))
    
    return points

def human_move_to(x, y):
    """Moves the mouse in a human-like way with dynamically adjusted points."""
    start_x, start_y = pyautogui.position()
    points = human_curve((start_x, start_y), (x, y))

    move_duration = random.uniform(*HUMAN_MOVE_DURATION)
    time_per_point = move_duration / len(points)

    for point in points:
        pyautogui.moveTo(*point, duration=0)
        time.sleep(time_per_point * random.uniform(0.1, 0.9))

        # Occasionally add tiny micro-movements
        if random.random() < MICRO_MOVEMENT_CHANCE:
            pyautogui.moveRel(
                random.uniform(-1, 1),
                random.uniform(-1, 1),
                duration=random.uniform(0.02, 0.05)
            )

def human_click():
    # Randomize click duration and add micro-movements
    click_duration =  random.uniform(0.3, 0.1)
    
    pyautogui.mouseDown()
    
    # Add small movements during click hold
    start_time = time.time()
    while time.time() - start_time < click_duration:
        time.sleep(0.02)
        pyautogui.moveRel(
            random.uniform(-0.5, 0.5),
            random.uniform(-0.5, 0.5),
            duration=0.01
        )
    
    pyautogui.mouseUp()

def parse_input_file(filename):
    events = []
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Input file {filename} not found!")
    
    with open(filename, 'r') as file:
        for line_num, line in enumerate(file, 1):
            line = line.strip()
            if not line:
                continue
            parts = [part.strip() for part in line.split(',')]
            if len(parts) != 4:
                print(f"Skipping invalid line {line_num}: {line}")
                continue
            try:
                x, y, action, delay = parts
                events.append((int(x), int(y), action, int(delay)))
            except ValueError:
                print(f"Skipping line {line_num} with invalid format: {line}")
    return events

def focus_game_window():
    try:
        # First, import pygetwindow
        try:
            import pygetwindow as gw
        except ImportError:
            print("pygetwindow is not installed. Please run: pip install pygetwindow")
            return False

        # Get all windows with a matching title
        game_windows = gw.getWindowsWithTitle(GAME_WINDOW_TITLE)
        
        if not game_windows:
            print(f"No window found with title: {GAME_WINDOW_TITLE}")
            print("Here are the available windows:")
            for window in gw.getAllWindows():
                if window.title.strip():  # Only show windows with titles
                    print(f"- {window.title}")
            return False
        
        # Focus the first matching window
        game_window = game_windows[0]
        if game_window.isMinimized:
            game_window.restore()  # Restore if minimized
        game_window.activate()
        time.sleep(0.5)  # Wait for window to focus
        return True
    except Exception as e:
        print(f"Error focusing game window: {e}")
        return False

def simulate_events(events, loops):
    stop_flag = False
    
    def set_stop_flag():
        nonlocal stop_flag
        stop_flag = True
    
    keyboard.add_hotkey(STOP_KEY, set_stop_flag)
    
    try:
        current_loop = 0
        while not stop_flag and (loops == 0 or current_loop < loops):
            current_loop += 1
            
            
            for event_num, event in enumerate(events, 1):
                if stop_flag:
                    print("\nStop requested by user!")
                    return
                
                if not focus_game_window():
                    print("Failed to focus game window!")
                    time.sleep(2)
                    return
                
                x, y, action, delay_ms = event
                
                # Human-like movement
                human_move_to(x, y)
                
                if action == 'Left Click':
                    human_click()
                
                # Randomized delay handling
               # base_delay = delay_ms / 1000
               # actual_delay = base_delay * random.uniform(0.9, 1.1)
               # delay_remaining = actual_delay
                
                #while delay_remaining > 0 and not stop_flag:
                #    step = min(0.1, delay_remaining)
                 #   time.sleep(step)
                  #  delay_remaining -= step
                    
                    # Random tiny movements during waiting
                   # if random.random() < 0.2:
                    #    pyautogui.moveRel(
                     #       random.uniform(-2, 2),
                      #      random.uniform(-2, 2),
                       #     duration=random.uniform(0.05, 0.1)
                        #)
    finally:
        keyboard.unhook_all()

def main():
    try:
        global CLICK_DELAY, GAME_WINDOW_TITLE
        
        print(f"Input file: {INPUT_FILE}")
        print(f"Stop key: {STOP_KEY.upper()}")
        print(f"Current click delay: {CLICK_DELAY} seconds")
        
        # Check for required modules
        try:
            import pygetwindow as gw
        except ImportError:
            print("pygetwindow is not installed. Please run: pip install pygetwindow")
            input("Press Enter to exit...")
            return
        
        # Automatically detect game window title if not set
        if GAME_WINDOW_TITLE == "Your Game Window":
            print("\nAvailable windows:")
            windows = [win for win in gw.getAllWindows() if win.title.strip()]
            for i, win in enumerate(windows, 1):
                print(f"{i}. {win.title}")
            
            try:
                choice = int(input("Enter the number of the game window: ")) - 1
                GAME_WINDOW_TITLE = windows[choice].title
                print(f"Selected window: {GAME_WINDOW_TITLE}")
            except (ValueError, IndexError):
                print("Invalid selection. Please manually set GAME_WINDOW_TITLE in the script.")
                input("Press Enter to exit...")
                return
        
        try:
            events = parse_input_file(INPUT_FILE)
        except FileNotFoundError as e:
            print(e)
            input("Press Enter to exit...")
            return
        
        if not events:
            print("No valid events found in the input file.")
            input("Press Enter to exit...")
            return
        
        try:
            loops = int(input("Enter number of loops (0 for infinite): "))
            new_delay = float(input(f"Enter click delay in seconds (current: {CLICK_DELAY}): "))
            CLICK_DELAY = max(0.05, new_delay)
        except ValueError:
            print("Invalid input. Using default values.")
            loops = 1
        
        print("\nStarting in 3 seconds... Ensure game window is visible!")
        for i in range(3, 0, -1):
            print(f"{i}...", end=' ')
            time.sleep(1)
        print("GO!")
        
        print(f"\nRunning Script, Current time: {time.ctime(time.time())}")
        simulate_events(events, loops)
        
    except Exception as e:
        print("\nAn error occurred:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        print("\nFull error traceback:")
        traceback.print_exc()
    
    finally:
        print("\nScript finished.")
        print(f"\nScript finished, Current time: {time.ctime(time.time())}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    try:
        # Check if required modules are installed
        missing_modules = []
        required_modules = ['pyautogui', 'keyboard', 'win32api', 'pygetwindow']
        
        for module in required_modules:
            try:
                __import__(module)
            except ImportError:
                missing_modules.append(module)
        
        if missing_modules:
            print("Missing required modules. Please install them using pip:")
            for module in missing_modules:
                print(f"pip install {module}")
            input("Press Enter to exit...")
            sys.exit(1)
            
        main()
    except Exception as e:
        print("\nCritical error occurred:")
        traceback.print_exc()
        input("Press Enter to exit...")