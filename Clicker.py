# Modules
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

# Settings
InputEventsFile = r"E:\AutoClicker\Clicks.txt"
StopKey = "esc"
AppTitle = "CROSSFIRE"

# Advanced
ClickDurationRange = (0.3, 1)
HumanMoveDuration = (0.3, 0.1) # Random movement duration range

# Human Curve
UseHumanCurve = False
BasePoints = 2 # Minimum number of points
MaxPoints = 10 # Maximum number of points for long distances

# DO NOT EDIT
JitterAmount = 0.7 # Reduce random jitter to prevent excessive slattering
MicroMovementChance = 0.03 # Less frequent micro-movements
MinCPU_Delay = 0.001
CPU_Time = 0.001
CheckRequiredModules = False

# Functions

# Check for admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# Check for Requirments
def CheckRequirements():
    import sys
    import subprocess
    import importlib.util

    required_modules = [
        'pyautogui',
        'keyboard',
        'pygetwindow',
        'win32api',  # part of pywin32
        'mouse'
    ]

    def is_module_installed(module_name):
        return importlib.util.find_spec(module_name) is not None

    def install_module(module_name):
        print(f"Installing {module_name}...")
        try:
            # Special case for pywin32 as it's imported as win32api
            if module_name == 'win32api':
                module_to_install = 'pywin32'
            else:
                module_to_install = module_name

            subprocess.check_call([sys.executable, '-m', 'pip', 'install', module_to_install])
            print(f"Successfully installed {module_name}")
            return True
        except subprocess.CalledProcessError as e:
            print(f"Failed to install {module_name}. Error: {e}")
            return False

    # First ensure pip is installed
    try:
        print("Checking pip installation...")
        subprocess.check_call([sys.executable, '-m', 'ensurepip', '--default-pip'])
       
    except subprocess.CalledProcessError as e:
        print(f"Failed to ensure pip installation. Error: {e}")
        return False

    # Check and install each required module
    all_modules_installed = True
    for module in required_modules:
        if not is_module_installed(module):
            print(f"{module} is not installed.")
            if not install_module(module):
                all_modules_installed = False    
    
    # End
    if all_modules_installed:
        print("\nAll required modules are installed successfully!")
        return True
    else:
        print("\nSome modules failed to install. Please install them manually:")
        print("pip install pyautogui keyboard pygetwindow pywin32 mouse")
        return False

# Mouse Movement
def MoveMouse(TargetPosition):

    def CalculatePointsNumber(Start, End):    
      Distance = math.dist(Start, End)
      return min(MaxPoints, max(BasePoints, int(Distance / 10)))  # Scale based on distance

    def HumanCurve(Start, End):
     """Generates a smooth Bezier curve path with fewer points for small distances."""
    
     # Get Points Count
     PointsCount = CalculatePointsNumber(Start, End)

     cp = [
        (Start[0] + (End[0] - Start[0]) * 0.25, Start[1] + (End[1] - Start[1]) * 0.25),
        (Start[0] + (End[0] - Start[0]) * 0.75, Start[1] + (End[1] - Start[1]) * 0.75)
     ]

     points = []
     for t in [i/PointsCount for i in range(PointsCount+1)]:
        x = (1-t)**3 * Start[0] + 3*(1-t)**2*t*cp[0][0] + 3*(1-t)*t**2*cp[1][0] + t**3*End[0]
        y = (1-t)**3 * Start[1] + 3*(1-t)**2*t*cp[0][1] + 3*(1-t)*t**2*cp[1][1] + t**3*End[1]

        points.append((x + random.uniform(-JitterAmount, JitterAmount),
                       y + random.uniform(-JitterAmount, JitterAmount)))
    
     return points

    MoveDuration = random.uniform(*HumanMoveDuration) -CPU_Time
    
    if UseHumanCurve:
     # Get Points that the mouse will go along them to the target  
     Points = HumanCurve(pyautogui.position(), TargetPosition)   
     TimePerPoint = (MoveDuration / len(Points))-CPU_Time
     for Point in Points:
        pyautogui.moveTo(*Point, duration=TimePerPoint)
     return
    else :
        pyautogui.moveTo(*TargetPosition, duration= MoveDuration)  
        return

# Mouse Input 
def MouseClick():
    # Random mouse click duration
    ClickDuration = random.uniform(*ClickDurationRange) -CPU_Time
    # Mouse down
    pyautogui.mouseDown()

    # Add small movements during click hold
    StartTime = time.time()
    while (time.time() - StartTime) < ClickDuration:
               
        pyautogui.moveRel(
            random.uniform(-0.5, 0.5),
            random.uniform(-0.5, 0.5),
            duration=0.01
        )
        time.sleep(MinCPU_Delay)
    
    pyautogui.mouseUp()
  
# Focus on App Window
def FocusOnWindow():
       
    try:
        # First, import pygetwindow
        try:
            import pygetwindow as gw
        except ImportError:
            print("pygetwindow is not installed. Please run: pip install pygetwindow")
            return False
        
        if not AppWindow :
            print(f"Invalid App Window")
            return False
        
        if AppWindow.isMinimized:
            AppWindow.restore()
        AppWindow.activate()
        time.sleep(0.5)
        return True

    except Exception as e:
        print(f"Error focusing app window: {e}")
        return False

def RefreshAppWindow():
      
    global AppWindow
    AppWindow = None # Reset AppWindow at start

    # Check for required modules
    try:
        import pygetwindow as gw
    except ImportError:
        print("pygetwindow is not installed. Please run: pip install pygetwindow")
        input("Press Enter to exit...")
        return
    
    # Get all windows and try to find a matching title
    FoundWindows = gw.getAllTitles()
      
    for WindowTitle in FoundWindows:
        if WindowTitle == AppTitle:
            windows = gw.getWindowsWithTitle(WindowTitle)
            if windows:  # Check if any windows were found
                AppWindow = windows[0]  # Store single window object, not list
                print(f"Found window: {WindowTitle}")
                return True
    
    # If couldn't find a window by provided title, give the user options of all available windwos
    if not AppWindow:
         # Print all windows
         for i, title in enumerate(FoundWindows, 1):
          print(f"{i}. {title}")

         # Try to get an input of user to select a window
         try:
          Selection = int(input("Enter the number of the app window: ")) -1
          if 0 <= Selection < len(FoundWindows):
              SelectedTitle = FoundWindows[Selection]
              windows = gw.getWindowsWithTitle(SelectedTitle)
              if windows:
                AppWindow = windows[0]  # Store single window object
                print(f"Selected window: {SelectedTitle}")
                return True
          else:
            raise IndexError

         except (ValueError, IndexError):
          print("Invalid selection. Please manually set AppTitle in the script.")
          input("Press Enter to exit...")
          return    
         
    # Check that app window is valid   
    if not AppWindow:
         print("Invalid App Window, please try again")
         input("Press Enter to exit...")

# Search for Input events
def GetInputEventsFromFile(filename):
    # Construct events variable
    events = []

    # Check if provided path is valid
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Input file {filename} not found!")
    
    # Open file
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

    # Return found events            
    return events

# Inputs Simulation
def SimulateInputEvents(Events, Loops):
   
   # Loop broken flag
   LoopBroken = False

   # Break the loop
   def BreakInputSimulation():
      nonlocal LoopBroken 
      LoopBroken = True
      print("\nStop requested by user!")

   keyboard.add_hotkey(StopKey, BreakInputSimulation)

   # Start the simulation Loop
   try:
      # innitialize loop count
      CurrentLoop = 0
      # start the loop
      while not LoopBroken and (Loops == 0 or CurrentLoop < Loops):       
         
         CurrentLoop += 1
         print(f"\nStarting loop {CurrentLoop}")

         for EventCount, Event in enumerate(Events, 1):
                if LoopBroken:                   
                    return
                
                if not FocusOnWindow():
                    print("Failed to focus app window! Retrying in 2 seconds...")
                    time.sleep(2)
                    continue

                XPosition, YPosition, Action, Duration= Event
                #print(f"Event {EventCount} : Moving to ({XPosition}, {YPosition} - {Action})", end = '\r')
                
                # Move Mouse
                MoveMouse((XPosition, YPosition))

                # Mouse Action
                if Action == 'Left Click':
                    MouseClick()              

   # On the end unhook stop key.
   finally:
    keyboard.unhook_all()

# Main

def main():

    if CheckRequiredModules:
       CheckRequirements()

    try:
      # Provide settings data
      print(f"\nInput events file: {InputEventsFile}")
      print(f"Stop Key: {StopKey.upper()}")

      # Try to ge the events from the given file
      try:
         InputEvents = GetInputEventsFromFile(InputEventsFile)
      except FileNotFoundError as e:
         print(e)
         input("Press Enter to exit...")
         return
      
      # Try to refresh the app window

      RefreshAppWindow()
      
      # Get Loops count
      try:
         Loops = int(input("\nEnter number of loops (0 for infinite): "))
      except ValueError:
         print("Invalid input. Using default values.")
         Loops = 5

     # Start a timer for 3 seconds before starting the loop
      print("\nStarting in 3 seconds... Ensure app window is visible!")
      for i in range(3, 0, -1):
       print(f"{i}...")
       time.sleep(1)
      print("GO!")

      print(f"\nRunning Script, Current time: {time.ctime(time.time())}")
      SimulateInputEvents(InputEvents, Loops)

    # if any error happened
    except Exception as e:
        print("\nAn error occurred:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        print("\nFull error traceback:")
        traceback.print_exc()
    
    # Ending
    finally:
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