import keyboard
import time

tracking_state = False
keys_to_count = {'ctrl', 'alt', 'f12'}
held_keys = set()
last_press_time = 0

def on_key_press(e):
    global tracking_state, last_press_time
    
    if e.name in keys_to_count:
        held_keys.add(e.name)
        
        if len(held_keys) == len(keys_to_count):
            current_time = time.time()
            
            if current_time - last_press_time < 0.5:
                tracking_state = not tracking_state
                print("Portal activated!" if tracking_state else "Portal deactivated!")
                last_press_time = 0
            else:
                last_press_time = current_time
            
            held_keys.clear()

def on_key_release(e):
    held_keys.discard(e.name)

keyboard.on_press(on_key_press)
keyboard.on_release(on_key_release)

while True:
    time.sleep(0.1)