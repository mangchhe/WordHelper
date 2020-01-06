from openpyxl import load_workbook
from gtts import gTTS

def engTTS(days, ws, value):

    with open(days +'일차.mp3','wb') as f:

        for i in range(len(value)):
            print('{}/{} {}% 진행중... - {} 추가중'.format(i+1,len(value),round(i/len(value)*100,1),value[i]))

            tts = gTTS(text=value[i], lang='en')

            tts.write_to_fp(f)
            tts.write_to_fp(f)

    f.close()

    print(days + '일차.mp3' + ' 파일을 추가하였습니다.')

if __name__ == '__main__':

    while True:

        try:
            wb = load_workbook('WordDirection.xlsx', data_only= True)
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

        engTTS(days, ws, value)

