import docx2txt
import time
import keyboard
import pyautogui
import random

def split(word):
    return list(word)

def blocked():
    for i in range(150):
        keyboard.block_key(i)
    keyboard.block_key('esc')
    keyboard.unblock_key('esc')

def unblocked():
    for i in range(150):
        keyboard.unblock_key(i)
        
wait = [1,2,1,0.1,0.1,0.1,0.1,0.2,0.2,0.1]
normal = [0.008,0.008,0.008,0.008,0.008,0.008,0.008,0.5,0.008,0.008,0.008,0.05,0.008,0.008]
# Driver code
Doc = input("Please Enter Word Document Name Below (eg . Bike.docx) (This only support docx files): \n")
raw_word = docx2txt.process(Doc)
word = split(raw_word)
blocked()

for i in word:
    if i == '’':
        x = word.index(i)
        word = word[:x]+["'"]+word[x+1:]

for i in word:
    if keyboard.is_pressed('esc'):
        unblocked()
        time.sleep(2)
        keyboard.wait('esc')
        blocked()
    else:
        pass
    pyautogui.write(i)
    if i == " ":
        time.sleep(random.choice(wait))
    else:
        time.sleep(random.choice(normal))
    


