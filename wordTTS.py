from openpyxl import load_workbook
import threading
import keyboard
from gtts import gTTS
from textblob import Word
from textblob import TextBlob

def execute():

    while True:
        if keyboard.is_pressed('tab'):
            print('\n[프로그램 종료] 엔터 눌러주세요.')
            break

def execute2(t):

    while t.is_alive():
        print('종료 : tab키 , 제작 : 하주')
        excelFileName = input('엑셀 파일 입력하세요 : ')

        try:
            wb = load_workbook(excelFileName + '.xlsx', data_only= True)
        except Exception as err:
            print('맞지 않는 파일 이름')
            continue


        days = input('일차 입력 : ')

        try:
            ws = wb[days + '일차']
        except Exception as err:
            print('맞지 않는 일차')
            continue

        value = []

        for row in ws.rows:
            value.append(row[0].value)

        preSelect = int(input("1 : 철자조사 2 : 음성파일 만들기 : "))

        if preSelect == 1:
            spellInspect(value)

            resultSelect = int(input("1 : 최종결정 2 : 되돌리기 : "))

            if resultSelect == 1:
                engTTS(days, ws, value)
            elif resultSelect == 2:
                continue
        elif preSelect == 2:
            engTTS(days, ws, value)

def engTTS(days, ws, value):

    with open(days +'일차.mp3','wb') as f:

        for i in range(len(value)):
            print('{}/{} {}% 진행중... - {} 추가중'.format(i+1,len(value),round(i/len(value)*100,1),value[i]))

            tts = gTTS(text=value[i], lang='en')

            tts.write_to_fp(f)
            tts.write_to_fp(f)

    f.close()

    print(days + '일차.mp3' + ' 파일을 추가하였습니다.')

def spellInspect(value):

    for i in range(len(value)):
        word = Word(value[i])
        if (word.spellcheck()[0][0] != value[i] and word.count(' ') == 0) or (word.spellcheck()[0][1] == 0 and word.count(' ') == 0):
            print("번호 : {} 단어 : {} -> ".format(i+1, value[i]), end="")
            for j in range(len(word.spellcheck())):
                if j == len(word.spellcheck())-1:
                    print("{}".format(word.spellcheck()[j][0]))
                    break
                print("{}, ".format(word.spellcheck()[j][0]), end="")

if __name__ == '__main__':

    t = threading.Thread(target=execute)
    t2 = threading.Thread(target=execute2, args=(t,))

    t.start()
    t2.start()
