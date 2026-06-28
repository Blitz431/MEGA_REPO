import keyboard 
import time

tracking_state = False  # Variable to track the state (True or False)
keys_to_count = {'ctrl', 'alt', 'f12'}  # Keys to count to activate the portal
held_keys = set()  # Set to store keys currently held down
sequence = []  # List to store the sequence of keys pressed
last_key_time = 0  # Variable to track the time of the last key press

def on_key_press(e):
    global tracking_state, held_keys
    
    if e.name in keys_to_count:
        held_keys.add(e.name)
        
        if len(held_keys) == len(keys_to_count):
            sequence.extend(sorted(held_keys))
            
            if len(sequence) >= 2 * len(keys_to_count) and all(sequence[i] == sequence[i + len(keys_to_count)] for i in range(len(keys_to_count))):
                tracking_state = not tracking_state  # Toggle the state
                print("Portal activated!")  # Perform the action when the portal is active
                
                # Reset the sequence for next use
                sequence.clear()
            
            held_keys.clear()

def on_key_release(e):
    if e.name in keys_to_count:
        held_keys.discard(e.name)

keyboard.on_press(on_key_press)
keyboard.on_release(on_key_release)

while True:
    time.sleep(0.1)