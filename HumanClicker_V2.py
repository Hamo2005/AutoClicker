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

# Configuration
InputEventsFile = r"E:\AutoClicker\Clicks_V2.txt"
StopKey = "esc"

# Advanced
ClickDelay = (0.3, 0.1)
HumanMoveDuration = (0.01, 0.1) # Random movement duration range

# Human Curve
BasePoints = 2 # Minimum number of points
MaxPoints = 7 # Maximum number of points for long distances

# DO NOT EDIT
JitterAmount = 0.7  # Reduced to prevent excessive jitter
MicroMovementChance = 0.03  # Lower chance for micro-movements
MinCPU_Delay = 0.02
CPU_Time = 0.001

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

# Mouse Movement
def HumanMove(Target):
    """Moves the mouse in a human-like way with dynamically adjusted points."""
    def CalculatePointsNumber(Start, End):    
      """Dynamically adjust the number of points based on distance."""
      Distance = math.dist(Start, End)
      return min(MaxPoints, max(BasePoints, int(Distance / 10)))  # Scale based on distance

    # Bézier curve implementation
    def HumanCurve(Start, End):
     """Generates a smooth Bezier curve path with fewer points for small distances."""
    
     # Get Points Count
     PointsCount = CalculatePointsNumber(Start, End)

     ControlPoints = [
          (Start[0] + (End[0] - Start[0]) * random.uniform(0.2, 0.4), 
          Start[1] + (End[1] - Start[1]) * random.uniform(0.2, 0.4)),
          (Start[0] + (End[0] - Start[0]) * random.uniform(0.6, 0.8), 
          Start[1] + (End[1] - Start[1]) * random.uniform(0.6, 0.8))
         ]

     points = []
     for t in [i/PointsCount for i in range(PointsCount+1)]:
        x = (1-t)**3 * Start[0] + 3*(1-t)**2*t*ControlPoints[0][0] + 3*(1-t)*t**2*ControlPoints[1][0] + t**3*End[0]
        y = (1-t)**3 * Start[1] + 3*(1-t)**2*t*ControlPoints[0][1] + 3*(1-t)*t**2*ControlPoints[1][1] + t**3*End[1]

        points.append((x + random.uniform(-JitterAmount, JitterAmount),
                       y + random.uniform(-JitterAmount, JitterAmount)))
    
     return points

    # Move like a human so no detection happens  
    Points = HumanCurve(pyautogui.position(), Target) 

    MoveDuration = random.uniform(*HumanMoveDuration)-CPU_Time  
    TimePerPoint = max(0.001, (MoveDuration / len(Points)) - CPU_Time)

    for Point in Points:
      pyautogui.moveTo(*Point, duration=0)    
      time.sleep(TimePerPoint* random.uniform(0.1, 0.9))

      # Occasionally add tiny micro-movements
      if random.random() < MicroMovementChance:
            pyautogui.moveRel(
                random.uniform(-1, 1),
                random.uniform(-1, 1),
                duration=random.uniform(0.02, 0.05)
            )

# Mouse Input 
def HumanClick(Target):
    # Random mouse click duration
    ClickDuration = random.uniform(*ClickDelay) - CPU_Time

    # Mouse down
    pyautogui.mouseDown()

    # Add small movements during click hold
    StartTime = time.time()
    while (time.time() - StartTime) < ClickDuration:
        time.sleep(MinCPU_Delay)   
        pyautogui.moveRel(
            random.uniform(-0.5, 0.5),
            random.uniform(-0.5, 0.5),
            duration=0.01
        )
        
    
    pyautogui.mouseUp()
 
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
                x, y, action, ActionDelay = parts
                events.append((int(x), int(y), action, int(ActionDelay)))
            except ValueError:
                print(f"Skipping line {line_num} with invalid format: {line}")

    # Return found events            
    return events

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
    FoundWindows = [win for win in gw.getAllWindows() if win.title.strip()]

    for i, win in enumerate(FoundWindows, 1):
      print(f"{i}. {win.title}") 
    
    # Try to get an input of user to select a window
    try:
          Selection = int(input("Enter the number of the app window: ")) -1
          if 0 <= Selection < len(FoundWindows):
              SelectedTitle = FoundWindows[Selection].title
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

# Inputs Simulation
def SimulateInputEvents(Events, Loops):
   
   # Loop broken flag
   LoopBroken = False
   
   def BreakLoop():
      nonlocal LoopBroken 
      LoopBroken = True
      print("\n Stopping Now")
   
   keyboard.add_hotkey(StopKey, BreakLoop)

   # Start the simulation Loop
   try:
      CurrentLoop = 0
      while not LoopBroken and (Loops == 0 or CurrentLoop < Loops):                
         CurrentLoop += 1    
         print(f"Current Loop: {CurrentLoop}")         

         for Event in Events:
                if LoopBroken:   
                    print("\nStopped by user request")                
                    return
                
                if not FocusOnWindow():
                    print("Failed to focus app window!")
                    return

                X, Y, Action, ActionDelay= Event

                # Screen Size
                ScreenWidth, ScreenHeight = pyautogui.size()
                ScreenWidth //= 2  # Divide by 2 to get center
                ScreenHeight //= 2

                # Compute target position
                TargetPosition = (ScreenWidth + X, ScreenHeight + Y)
                               
                # Move Mouse
                HumanMove(TargetPosition)
                if Action == 'Left Click':
                  HumanClick(TargetPosition)   

                DelayForNextInput = (ActionDelay /1000) * random.uniform(0.9, 1.1)

                StartTime = time.time()
                while (time.time() - StartTime) < DelayForNextInput:
                    time.sleep(min(0.1, DelayForNextInput))  
                    
                    # Random tiny movements during waiting
                    if random.random() < 0.2:
                        pyautogui.moveRel(
                            random.uniform(-2, 2),
                            random.uniform(-2, 2),
                            duration=random.uniform(0.05, 0.1)
                        )         

   # On the end unhook stop key.
   finally:
    keyboard.unhook_all()

# Main

def main():
    try:
      print("IMPORTANT: Don't make a pattern that keeps the mouse in same place for too long, so the code doesn't get detected!")
      
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