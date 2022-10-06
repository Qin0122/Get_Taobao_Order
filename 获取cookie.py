from selenium import webdriver
from time import sleep
import json

if __name__ == '__main__':
  driver = webdriver.Chrome()
  driver.maximize_window()
  driver.get('https://login.taobao.com/member/login.jhtml?')
  sleep(20)

  dictCookies = driver.get_cookies() # 获取list的cookies
  jsonCookies = json.dumps(dictCookies) # 转换成字符串保存
  with open('taobao_cookies.txt', 'w') as f:
    f.write(jsonCookies)
  print('cookies保存成功！')