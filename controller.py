import cv2
import mediapipe as mp
import pyautogui
import time
import subprocess
import os
import pygetwindow as gw
import win32gui
import win32con

# Helper: Keep window always on top
def make_window_always_on_top(window_name):
    hwnd = win32gui.FindWindow(None, window_name)
    if hwnd:
        win32gui.SetWindowPos(
            hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
        )

# 1. Open PowerPoint in slideshow mode
ppt_path = r"C:\Users\Shreya\Desktop\AICTE\AI_Presentation_Controller\Water Born Diseases.pptx"
if not os.path.exists(ppt_path):
    print(f"Error: PPT file not found at {ppt_path}")
    exit(1)

print("Opening PowerPoint slideshow...")
subprocess.Popen(['start', 'powerpnt', '/s', ppt_path], shell=True)
time.sleep(5)  # wait for PPT to open

# 2. Focus PowerPoint
def focus_powerpoint():
    windows = gw.getWindowsWithTitle('PowerPoint Slide Show')
    if not windows:
        windows = gw.getWindowsWithTitle('PowerPoint')
    if windows:
        ppt_win = windows[0]
        ppt_win.activate()
        time.sleep(0.2)
        return True
    return False

focus_powerpoint()

# 3. Setup MediaPipe Hands
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(max_num_hands=1)
mp_draw = mp.solutions.drawing_utils

# 4. Open webcam
cap = cv2.VideoCapture(0)

# 5. Finger tip IDs in MediaPipe
tip_ids = [4, 8, 12, 16, 20]
pointer_mode = False
last_action_time = 0
cooldown = 1.5

print("Controller started. Show 1 finger: Previous Slide, 2 fingers: Next Slide, 5 fingers: Toggle Pointer Mode.")
print("Press 'q' to quit.")

first_show = True  # track first time to make window always on top

while True:
    success, img = cap.read()
    if not success:
        continue

    img = cv2.flip(img, 1)
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = hands.process(img_rgb)

    lm_list = []
    if results.multi_hand_landmarks:
        hand_landmarks = results.multi_hand_landmarks[0]
        for id, lm in enumerate(hand_landmarks.landmark):
            h, w, _ = img.shape
            cx, cy = int(lm.x * w), int(lm.y * h)
            lm_list.append((id, cx, cy))

        mp_draw.draw_landmarks(img, hand_landmarks, mp_hands.HAND_CONNECTIONS)

        fingers = []
        # Thumb (right hand logic)
        if lm_list[tip_ids[0]][1] > lm_list[tip_ids[0] - 1][1]:
            fingers.append(1)
        else:
            fingers.append(0)
        # Other fingers
        for i in range(1, 5):
            if lm_list[tip_ids[i]][2] < lm_list[tip_ids[i] - 2][2]:
                fingers.append(1)
            else:
                fingers.append(0)

        total_fingers = sum(fingers)
        current_time = time.time()

        if current_time - last_action_time > cooldown:
            focus_powerpoint()
            if total_fingers == 1:
                pyautogui.press('left')
                print("⬅️ Previous Slide (1 finger)")
                last_action_time = current_time
            elif total_fingers == 2:
                pyautogui.press('right')
                print("➡️ Next Slide (2 fingers)")
                last_action_time = current_time
            elif total_fingers == 5:
                pointer_mode = not pointer_mode
                print(f"🖱️ Pointer Mode {'ON' if pointer_mode else 'OFF'}")
                last_action_time = current_time

        if pointer_mode:
            index_pos = (lm_list[8][1], lm_list[8][2])
            cv2.circle(img, index_pos, 15, (0, 0, 255), cv2.FILLED)

    cv2.putText(img, f"Pointer: {'ON' if pointer_mode else 'OFF'}", (10, 30),
                cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

    cv2.imshow("Presentation Controller", img)

    # Make webcam window always on top once
    if first_show:
        make_window_always_on_top("Presentation Controller")
        first_show = False

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
