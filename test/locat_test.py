import pyautogui
import time
img = 'test_red_square1.png'
while True:
    import pyscreeze

    # location = pyscreeze.locateCenterOnScreen(img)
    location = pyautogui.locateCenterOnScreen(img)
    print(location)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=2, interval=0.2, duration=0.2, button='left')
        break
    print("未找到匹配图片,0.1秒后重试")
    time.sleep(0.1)

