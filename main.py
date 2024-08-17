import os
import time
import pandas as pd

from log import log
from mode import one_of_one, lot_of_one, crop_img

if __name__ == '__main__':
  print('「照片貼手」\n'
        '版本：1.0.2_08141220\n'
        '作者：蔡智楷 C.K.SAI\n'
        '本軟體分享於「交通鴿手 trafficpigeon.com 」\n')
  print('功能介紹：\n'
        '本軟體可一次將多張圖片，進行裁切、排版，製作成照片黏貼表。\n')
  print('使用說明：\n'
        '一、請將圖片放入「Image」資料夾當中，建議預先更改檔名，排定排序。\n'
        '二、可於「Setting」的Excel檔中自訂預設的文字。\n'
        '三、執行前，請確定程式資料夾內的檔案皆已儲存並關閉。\n')
  input('請詳閱使用說明，輸入任意鍵繼續……')
  log().info('程式開始執行')

  # 讀取設定，並儲存成字典
  try:
    setting = pd.read_excel(os.path.abspath('Setting.xlsx'), sheet_name='預設文字', index_col=0)
    options = {
      'title': str(setting.at['標題', '預設文字']),
      'time': str(setting.at['攝影時間', '預設文字']),
      'place': str(setting.at['攝影地點', '預設文字']),
      'describe': str(setting.at['說明', '預設文字']),
    }
    # 如果選項為空，設為空字串
    for key,val in options.items():
      if val == 'nan':
        options[key] = ''
    log().debug(options)
  except:
    log().error('讀取「Setting」檔失敗，請確認檔案已關閉。', exc_info=True)

  time.sleep(0.5)

  while True:
    print('\n◎ 請選擇模式：'
          '\n(1) 照片黏貼（一格單張）'
          '\n(2) 照片黏貼（一格多張 - 適用於手機截圖)')
    mode = input('\n● 請輸入數字選擇：')
    log().debug(f'已輸入模式「{mode}」')

    if mode == '1':
      no_crop = input('● 是否執行長截圖分割？如不執行，請輸入「 N 」：')
      if no_crop.upper() != 'N':
        crop_img()
        time.sleep(0.5)
      one_of_one(options)
      time.sleep(0.5)
      input('輸入任意鍵退出程式……')
      break

    elif mode == '2':
      no_crop = input('● 是否執行長截圖分割？如不執行，請輸入「 N 」：')
      if no_crop.upper() != 'N':
        crop_img()
        time.sleep(0.5)
      lot_of_one(options)
      time.sleep(0.5)
      input('輸入任意鍵退出程式……')
      break

    else:
      log().error('模式輸入錯誤，請重新輸入！')
