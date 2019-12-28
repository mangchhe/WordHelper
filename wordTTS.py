from openpyxl import load_workbook
import threading
import keyboard
from gtts import gTTS

def execute():

    while True:
        if keyboard.is_pressed('tab'):
            print('\n[프로그램 종료] 엔터 눌러주세요.')
            break

def execute2(t):

    while t.is_alive():
        print('종료 : tab키 , 제작 : 하주현')
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

        engTTS(days, ws)

def engTTS(days, ws):

    value = []

    for row in ws.rows:
        value.append(row[0].value)

    with open(days +'일차.mp3','wb') as f:

        for i in range(len(value)):
            print('{}/{} {}% 진행중... - {} 추가중'.format(i,len(value),round(i/len(value)*100,1),value[i]))

            tts = gTTS(text=value[i], lang='en')

            tts.write_to_fp(f)
            tts.write_to_fp(f)

    f.close()

if __name__ == '__main__':

    t = threading.Thread(target=execute)
    t2 = threading.Thread(target=execute2, args=(t,))

    t.start()
    t2.start()
