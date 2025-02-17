# Modules
import pyautogui
import time
import sys


# Settings
MinCPU_Delay = 0.001

# Record Mouse Position
def main():
      ScreenWidth, ScreenHeight = pyautogui.size()
      CenterX, CenterY = ScreenWidth // 2, ScreenHeight // 2     
      while True:
          x, y = pyautogui.position()
          RelativeX = x - CenterX
          RelativeY = y - CenterY
          sys.stdout.write(f"\rRelative Mouse Position = ({RelativeX}, {RelativeY}) | Current Position = ({x}, {y})")
          sys.stdout.flush()
          time.sleep(MinCPU_Delay)  # Reduce CPU usage


if __name__ == "__main__":

    main()
     