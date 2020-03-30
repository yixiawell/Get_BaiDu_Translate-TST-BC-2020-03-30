import requests, xlrd, os
from playsound import playsound
from urllib import parse
'''
作用：先读取excel中sheet1中的单词
    再读取voice文件夹中的文件列表，并比对是否存在，若存在直接播放，不存在下载完成后再播放
'''
#读取excel
def read_excel():
    '''
    :return: 需要处理的单词
    '''
    if not os.path.exists(r"C:\Users\ASUS-PC\Desktop\vido"):
        os.mkdir(r"C:\Users\ASUS-PC\Desktop\vido")
    book = xlrd.open_workbook(r"C:\Users\ASUS-PC\Desktop\单词记忆.xlsm")
    sheet1 = book.sheets()[0]
    word = sheet1.cell(0, 0).value.strip()
    return word

#打开文件列表并比对
def open_file(word):
    file_list = os.listdir(r"C:\Users\ASUS-PC\Desktop\vido")
    word_list = [i.replace(".mp3", "") for i in file_list]
    print(word_list,file_list)
    if word in word_list:
        #调用播放音频
        play_vido("{}.mp3".format(word))
    else:
        #调用获取音频
        get_vido(word)

#播放音频
def play_vido(files):
    filename = r"C:\Users\ASUS-PC\Desktop\vido\{}".format(files)
    print(filename)
    playsound(filename)

#获取音频
def get_vido(word):
    word1 = parse.quote(word)
    print(word1)
    url = "https://fanyi.baidu.com/gettts?lan=en&text={}&spd=3&source=web".format(word1)
    req = requests.get(url)
    with open(r"C:\Users\ASUS-PC\Desktop\vido\{}.mp3".format(word), "wb+") as f:
        f.write(req.content)
    play_vido("{}.mp3".format(word))

if __name__ == '__main__':
    word = read_excel()
    open_file(word)
    # get_vido("recoverable")