import sys, os, shutil, shelve, re, datetime, openpyxl, json, random

from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QApplication, QMessageBox, QLabel, QLineEdit, QTextEdit, QTableView, QVBoxLayout, QTabWidget, QInputDialog
from PyQt5.QtGui import QFont, QIcon, QPixmap, QStandardItemModel, QStandardItem
from PyQt5.QtCore import QCoreApplication, pyqtSlot

    #グローバル変数の定義、初期化
TestWordList = [["" for column in range(4)] for row in range(10)]
JapEngWordIdiom_mode = ""
Number_RecentTestFile = 0
SelectedRecentFilenameIndex = 0
ExistingRecentTestFilename = ["" for column in range(5)]
Original_Excel_Max_Row_Number = 0
Max_Row = {'EngWord':0, 'EngIdiom':0, 'JapWord':0}
mode_filename_dic = {'EngWord':'EngJapWordFile', 'EngIdiom':'EngJapIdiomFile', 'JapWord':'JapEngWordFile'}
mode_message_dic = {'EngWord':'英-->日（単語）は', 'EngIdiom':'英-->日（イディオム）は', 'JapWord':'日-->英（単語）は'}
Question_Number = 10
Memorized_count_in10 = 0
ShowOneByOnePassTime = 0

class StartMenu(QWidget):
  def __init__(self, parent):
    super().__init__(parent)
    self.master = parent

    self.label = QLabel(self)
    pixmap = QPixmap('AtFacebook.jpg')
    self.label.setPixmap(pixmap)
    self.label.move(410, 40)

    self.label2 = QLabel(self)
    self.label2.setText("Copyright 2018 Hiromi Maeda")
    self.label2.move(900,840)

    self.button1 = QPushButton('①英語　--> 日本語＜単語＞', self)
    self.button1.move(280, 230)
    self.button1.clicked.connect(self.Eng_Jap_Word)

    self.button2 = QPushButton('②英語　--> 日本語＜イディオム＞　', self)
    self.button2.move(550, 230)
    self.button2.clicked.connect(self.Eng_Jap_Idiom)

    self.button3 = QPushButton('③日本語 --> 英語　　　　　　　　　　', self)
    self.button3.move (400, 270)
    self.button3.clicked.connect(self.Jap_Eng)

    self.button4 = QPushButton('④前回、途中中断した１０問の再テスト', self)
    self.button4.move (400, 310)
    self.button4.clicked.connect(self.RestartFromCancel)

    self.button5 = QPushButton('⑤前回テストの再テスト　　　　　　　　　', self)
    self.button5.move (400, 350)
    self.button5.clicked.connect(self.RedoLastTest)

    self.button6 = QPushButton('⑥前回テストの「日<-->英」逆テスト　　', self)
    self.button6.move (400, 390)
    self.button6.clicked.connect(self.ReverseTest)

    self.button6 = QPushButton('⑦ファイルから新しい単語の追加（未完成）', self)
    self.button6.move (400, 430)
    self.button6.clicked.connect(self.AddNewWord)


    self.quitbutton1 = QPushButton('⑧終了します', self)
    self.quitbutton1.move (460, 485)
    self.quitbutton1.clicked.connect(QCoreApplication.instance().quit)

    self.label3 = QLabel(self)
    self.label3.setText("********* あなたの今までの成績 *********\n１：　英-->日（単語）       計 \n　　　現在までにテストした単語数：\n　　　覚える事の出来た単語数：")
    self.label3.move(370,520)

    flag = self.ShowPerformance('EngWord')
    if flag == 0:
      self.label3 = QLabel(self)
      self.label3.setText("********* あなたの今までの成績 *********\n１：英-->日（単語）       計 \n    今までにテスト未実施です：\n     　　　　　　　　　　　　")
      self.label3.move(370,520)

    self.label4 = QLabel(self)
    self.label4.setText("********* あなたの今までの成績 *********\n２：　英-->日（イディオム） 計\n　　　現在までにテストした単語数：\n　　　覚える事の出来た単語数：　　")
    self.label4.move(370,620)

    flag = self.ShowPerformance('EngIdiom')
    if flag == 0:
      self.label4 = QLabel(self)
      self.label4.setText("********* あなたの今までの成績 *********\n２：　英-->日（単語）       計 \n    今までにテスト未実施です：\n     　　　　　　　　　　　　")
      self.label4.move(370,620)

    self.label5 = QLabel(self)
    self.label5.setText("********* あなたの今までの成績 *********\n３：　日-->英              計\n　　　現在までにテストした単語数：\n　　　覚える事の出来た単語数：　　")
    self.label5.move(370,720)

    flag = self.ShowPerformance('JapWord')
    if flag == 0:
      self.label5 = QLabel(self)
      self.label5.setText("********* あなたの今までの成績 *********\n３：　日-->英              計\n    今までにテスト未実施です：\n     　　　　　　　　　　　　")
      self.label5.move(370,720)





  def ShowPerformance(self, mode0):
    mode = mode0
    global Max_Row
    global mode_filename_dic

    y = os.path.exists("./"+mode_filename_dic[mode]+".dat")
    if y == False:
      return 0
    else:
      shelf_file = shelve.open(mode_filename_dic[mode])

    #if mode == 'EngWord':
      #y = os.path.exists("./EngJapWordFile.dat")
      #if y == False:
        #return 0
      #else
        #shelf_file = shelve.open('EngJapWordFile')
    #elif mode == 'EngIdiom':
      #y = os.path.exists("./EngJapIdiomFile.dat")
      #if y == False:
        #return 0
      #else
        #shelf_file = shelve.open('EngJapIdiomFile')
    #elif mode == 'JapWord':
      #y = os.path.exists("./JapEngWordFile.dat")
      #if y == False:
        #return 0
      #else
        #shelf_file = shelve.open('JapEngWordFile')

    temp_word_array = shelf_file['word']
    Max_Row[mode] = shelf_file['max_row']
    shelf_file.close()

    counter1 = 0
    counter2 = 0

    for row_num in range(Max_Row[mode]):
      if temp_word_array[row_num][2] == 'Y': #tested?
        counter1 += 1
        if temp_word_array[row_num][4] == 'Y': #memorized?
          counter2 += 1

    if mode == 'EngWord':
      self.label6_1 = QLabel(self)
      self.label6_1.setText(str(Max_Row[mode]))
      self.label6_1.move(600,535)
      self.label6_2 = QLabel(self)
      self.label6_2.setText(str(counter1))
      self.label6_2.move(600,555)
      self.label6_3 = QLabel(self)
      self.label6_3.setText(str(counter2))
      self.label6_3.move(600,575)
    elif mode == 'EngIdiom':
      self.label7_1 = QLabel(self)
      self.label7_1.setText(str(Max_Row[mode]))
      self.label7_1.move(600,635)
      self.label7_2 = QLabel(self)
      self.label7_2.setText(str(counter1))
      self.label7_2.move(600,655)
      self.label7_3 = QLabel(self)
      self.label7_3.setText(str(counter2))
      self.label7_3.move(600,675)
    elif mode == 'JapWord':
      self.label8_1 = QLabel(self)
      self.label8_1.setText(str(Max_Row[mode]))
      self.label8_1.move(600,735)
      self.label8_2 = QLabel(self)
      self.label8_2.setText(str(counter1))
      self.label8_2.move(600,755)
      self.label8_3 = QLabel(self)
      self.label8_3.setText(str(counter2))
      self.label8_3.move(600,775)


  def Eng_Jap_Word(self):
    global JapEngWordIdiom_mode

    #self.All_Erase1()
    JapEngWordIdiom_mode = 'EngWord'
    self.master.setCurrentIndex(1)


  def Eng_Jap_Idiom(self):
    global JapEngWordIdiom_mode

    #self.All_Erase1()
    JapEngWordIdiom_mode = 'EngIdiom'
    self.master.setCurrentIndex(1)


  def Jap_Eng(self):
    global JapEngWordIdiom_mode

    #self.All_Erase1()
    JapEngWordIdiom_mode = 'JapWord'
    self.master.setCurrentIndex(1)


  def RestartFromCancel(self):
    global TestWordList
    global mode_message_dic
    global Question_Number
    global Memorized_count_in10

    #self.All_Erase1()

    y = os.path.exists("./10WordLogForCancel.dat")
    if y == False:
      msgBox = QMessageBox.warning(self, '警告', "前回中断(Abort)で終了した形跡は無いです、別の選択肢を選んで下さい", QMessageBox.Ok)
      return

    filename = '10WordLogForCancel'
    shelf_file = shelve.open(filename)
    TestWordList = shelf_file['word']
    mode = shelf_file['mode']

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1

    Question_Number = number
    if number < 10:
      message1 = "残りの問題数が"+number+"問になりました"
      msgBox = QMessageBox.information(self, '通知', message1, QMessageBox.Ok)

    self.master.setCurrentIndex(2)
    #Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    if (Question_Number < 10) and (Memorized_count == Question_Number):
      message2 = "おめでとうございます"+mode_message_dic[mode]+"全問覚えました"
      msgBox = QMessageBox.information(self, '通知', message2, QMessageBox.Ok)


    if Memorized_count_in10 == -1:
      #self.All_Erase3()
      shelf_file.close()
      return

    #self.All_Erase3()

    self.master.setCurrentIndex(3)
    #self.ShowResult()

    msgBox3 = QMessageBox()
    msgBox3.setText("結果をファイルにSaveしますか？")
    msgBox3.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msgBox3.setDefaultButton(QMessageBox.Yes)
    screenGeometry3 = QApplication.desktop().availableGeometry()
    screenGeo3 = screenGeometry3.bottomRight()
    msgGeo3 = msgBox3.frameGeometry()
    msgGeo3.moveBottomRight(screenGeo3)
    msgBox3.show()
    ret3 = msgBox3.exec_()

    if ret3 == QMessageBox.Yes:
      #１０問第をファイルにSave、Historyファイルの更新
      self.Write10WordTest()

      #Master File更新（Indexにより）
      self.UpdateFiles()

      os.unlink("./10WordLogForCancel.bak")
      os.unlink("./10WordLogForCancel.dat")
      os.unlink("./10WordLogForCancel.dir")

      return

    elif ret3 == QMessageBox.No:
      return




  def RedoLastTest(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Question_Number
    global Memorized_count_in10

    #self.All_Erase1()

    y = os.path.exists("./LastTest.dat")
    if y == False:
      msgBox0 = QMessageBox.warning(self, '警告', "前回のテストは見当たりません、別の選択肢を選んで下さい", QMessageBox.Ok)
      return
    else:
      filename = 'LastTest'
      shelf_file = shelve.open(filename)
      TestWordList = shelf_file['word']
      print(TestWordList)
      todaydatetime = shelf_file['date']
      accuracy = shelf_file['accuracy_rate']
      mode = shelf_file['mode']
      shelf_file.close()


    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1
    Question_Number = number
    if number <10:
      message1 = "残りの問題数が"+number+"問になりました"
      msgBox = QMessageBox.information(self, '通知', message1, QMessageBox.Ok)

    Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    if Memorized_count_in10 == -1: #途中Cancelボタンが押された
      #今回未実施だった１０問を、次回の為にSaveする
      y = os.path.exists("./10WordLogForCancel.dat")
      if y == True:
        print(y)
        #os.unlink("./10WordLogForCancel.bak")
        #os.unlink("./10WordLogForCancel.dat")
        #os.unlink("./10WordLogForCancel.dir")

      filename = '10WordLogForCancel'
      shelf_file1 = shelve.open(filename)
      shelf_file1['word'] = TestWordList
      shelf_file1['date'] = tpdaydatetime
      shelf_file1['accuracy_rate'] = accuracy
      shelf_file1['mode'] = mode
      shelf_file1.close()
      return

    #self.All_Erase3()

    self.ShowResult(Memorized_count_in10)

    msgBox3 = QMessageBox()
    msgBox3.setText("結果をファイルにSaveしますか？")
    msgBox3.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msgBox3.setDefaultButton(QMessageBox.Yes)
    screenGeometry3 = QApplication.desktop().availableGeometry()
    screenGeo3 = screenGeometry3.bottomRight()
    msgGeo3 = msgBox3.frameGeometry()
    msgGeo3.moveBottomRight(screenGeo3)
    msgBox3.show()
    ret3 = msgBox3.exec_()

    JapEngWordIdiom_mode = mode
    if ret3 == QMessageBox.Yes:
      #１０問第をファイルにSave、Historyファイルの更新
      self.Update_LastTest_After_Redoing_Last10WordTest()#SimpleなWrite10WordTestを作る
                                    #Trial number, 正解?のみUpdate

      #Master File更新（Indexにより）
      self.UpdateFiles_After_Redoing_Last10WordTest(number) #SimpleなUpdateFiles()を作る
                            #Trial number, 正解?のみUpdate
      return

    elif ret3 == QMessageBox.No:
      return



  def ReverseTest(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Question_Number
    global Memorized_count_in10

    #self.All_Erase1()
    y = os.path.exists("./LastTest.dat")
    if y == False:
      msgBox0 = QMessageBox.warning(self, '警告', "前回のテストは見当たりません、別の選択肢を選んで下さい", QMessageBox.Ok)
      return
    else:
      filename = 'LastTest'
      shelf_file = shelve.open(filename)
      TestWordList = shelf_file['word']
      #print(self.TestWordList)
      todaydatetime = shelf_file['date']
      count = shelf_file['accuracy_rate']
      mode = shelf_file['mode']
      shelf_file.close()

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1
    Question_Number = number
    if number <10:
      message1 = "残りの問題数が"+number+"問になりました"
      msgBox = QMessageBox.information(self, '通知', message1, QMessageBox.Ok)

    for i in range(Question_Number):
      tempword = TestWordList[i][0]
      TestWordList[i][0] = self.TestWordList[i][1]
      TestWordList[i][1] = tempword
      TestWordList[i][2] = ''

    Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    if Memorized_count_in10 == -1: #途中Cancelボタンが押された
      #今回未実施だった１０問を、次回の為にSaveする
      y = os.path.exists("./10WordLogForCancel.dat")
      if y == True:
        print(y)
        #os.unlink("./10WordLogForCancel.bak")
        #os.unlink("./10WordLogForCancel.dat")
        #os.unlink("./10WordLogForCancel.dir")
      filename = '10WordLogForCancel'
      shelf_file = shelve.open(filename)
      shelf_file['word'] = TestWordList
      shelf_file.close()

    #self.All_Erase3()

    self.master.setCurrentIndex(3)
    #self.ShowResult(count)

    msgBox3 = QMessageBox()
    msgBox3.setText("結果をSaveしますか？")
    msgBox3.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msgBox3.setDefaultButton(QMessageBox.Yes)
    screenGeometry3 = QApplication.desktop().availableGeometry()
    screenGeo3 = screenGeometry3.bottomRight()
    msgGeo3 = msgBox3.frameGeometry()
    msgGeo3.moveBottomRight(screenGeo3)
    msgBox3.show()
    ret3 = msgBox3.exec_()

    #１０をファイルにSave、Historyファイルの更新
    JapEngWordIdiom_mode = mode
    self.Write10WordTest()

    #Master File更新（Indexにより）
    self.UpdateFiles()


  def AddNewWord(self):
    print("Under Construction")



class NextMenu(QWidget):
  def __init__(self, parent):
    super().__init__(parent)
    self.master = parent

    self.label = QLabel(self)
    pixmap = QPixmap('AtFacebook.jpg')
    self.label.setPixmap(pixmap)
    self.label.move(410, 40)

    self.label2 = QLabel(self)
    self.label2.setText("Copyright 2018 Hiromi Maeda")
    self.label2.move(900,840)



    self.label6 = QLabel(self)
    self.label6.setText("********* １０問の問題を出します *********")
    self.label6.move(390,200)


    self.button6 = QPushButton('①新規の単語テスト', self)
    self.button6.move (430, 230)
    self.button6.clicked.connect(self.NewWordTest)

    self.button7 = QPushButton('②最近の復習（再テスト）', self)
    self.button7.move (430, 280)
    self.button7.clicked.connect(self.RecentTest)

    self.button8 = QPushButton('③覚えきれていない単語の再テスト', self)
    self.button8.move (400, 330)
    self.button8.clicked.connect(self.NotMemorizedWordTest)

    self.button9 = QPushButton('④覚えたと思っている単語の再テスト', self)
    self.button9.move (400, 380)
    self.button9.clicked.connect(self.MemorizedWordTest)

    self.quitbutton2 = QPushButton('⑤終了します', self)
    self.quitbutton2.move (440, 430)
    self.quitbutton2.clicked.connect(self.GoBackToStartMenu)

    #self.button10 = QPushButton('上記の選択（番号入力）する為、このボタンをPush！', self)
    #self.button10.move(400, 550)
    #self.button10.hide()



  def NewWordTest(self):
    """
    役割 : 日英単語、日英イディオム、英日で、新規の単語を質問するモード
    引数 : 無
    戻り値  : 無
    """
    global TestWordList
    global Question_Number
    global Memorized_count_in10

    Not_Read_number = self.Extract10Words("New")
    #print(TestWordList)

    #self.All_Erase2()

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        #print(TestWordList[i][0], TestWordList[i][1])
        number += 1
    Question_Number = number

    self.master.setCurrentIndex(2)  #ShowOneByOne
    #Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    #self.All_Erase3()

    #self.master.setCurrentIndex(3)    #ShowResult
    #self.ShowResult(Memorized_count_in10)

    #self.master.setCurrentIndex(2)


  def Extract10Words(self, mode):
    """
    役割 : 日英単語、日英イディオム、英日ファイルから質問となる１０問を選ぶ　TestWordList(Global)に代入
    引数 : mode --> 日英単語、日英イディオム、英日)
    戻り値  : counter --> それぞれのmodeで、未試験、記憶、未記憶のものがいくつ残っているか
    """
    global TestWordList
    global JapEngWordIdiom_mode
    global Max_Row

    temp_mode = mode
    self.index = [0 for i in range(10)]

    if JapEngWordIdiom_mode == 'EngWord':
      shelf_file = shelve.open('EngJapWordFile')
    elif JapEngWordIdiom_mode == 'EngIdiom':
      shelf_file = shelve.open('EngJapIdiomFile')
    elif JapEngWordIdiom_mode == 'JapWord':
      shelf_file = shelve.open('JapEngWordFile')

    self.wordlist = shelf_file['word']
    max_row = shelf_file['max_row']
    self.lastdate = shelf_file['date']

    #condition = True
    self.BufferWordList = [["" for column in range(4)] for row in range(max_row)]
    counter = 0

    TestWordList = [["" for column in range(4)] for row in range(10)]
    for row_num in range(max_row):
      if mode == "New":
        if self.wordlist[row_num][2] == 'N':  #already_read
          self.BufferWordList[counter][0] = self.wordlist[row_num][0]
          self.BufferWordList[counter][1] = self.wordlist[row_num][1]
          counter += 1
      elif mode == "NotMemorized":
        if self.wordlist[row_num][4] == 'N':  #memorized
          self.BufferWordList[counter][0] = self.wordlist[row_num][0]
          self.BufferWordList[counter][1] = self.wordlist[row_num][1]
          counter += 1
      elif mode == "Memorized":
        if self.wordlist[row_num][4] == 'Y':  #memorized
          self.BufferWordList[counter][0] = self.wordlist[row_num][0]
          self.BufferWordList[counter][1] = self.wordlist[row_num][1]
          counter += 1

    if counter > 50:
      for i in range(10):
        self.index = random.randint(0, counter-1)
        TestWordList[i][0] = self.BufferWordList[self.index][0]
        TestWordList[i][1] = self.BufferWordList[self.index][1]
        TestWordList[i][3] = self.index
        if mode == "Memorized":
          TestWordList[i][2] = 'Y'
        else:
          TestWordList[i][2] = 'N'

        #print(TestWordList[i][0])
        #print(TestWordList[i][1])
        #print(TestWordList[i][2])
    elif (counter >=10) and (counter <= 50) :
      for i in range(10):
        TestWordList[i][0] = self.BufferWordList[i][0]
        TestWordList[i][1] = self.BufferWordList[i][1]
        TestWordList[i][3] = self.index
        if mode == "Memorized":
          TestWordList[i][2] = 'Y'
        else:
          TestWordList[i][2] = 'N'

    elif (counter < 10):
      for i in range(counter):
        TestWordList[i][0] = self.BufferWordList[i][0]
        TestWordList[i][1] = self.BufferWordList[i][1]
        TestWordList[i][3] = self.index
        if mode == "Memorized":
          TestWordList[i][2] = 'Y'
        else:
          TestWordList[i][2] = 'N'

    shelf_file.close()

    return counter
    #このcounterは、modeがNewの場合、テスト未実施の数、
    #modeがMemorizedの場合、Memorizeした単語の数、
    #modeがNotMemorizedの場合、Memorizeしていない単語の数


  def RecentTest(self):

    #self.All_Erase2()

    self.SelectRecentFile()
    #self.All_Erase4()
    #self.TestSelectedRecentFile()
    #self.All_Erase5()





  def NotMemorizedWordTest(self):
    global TestWordList
    global Question_Number
    global Memorized_count_in10

    Not_Memorized_number = self.Extract10Words("NotMemorized")

    if Not_Memorized_number ==0:
      msgBox0 = QMessageBox.warning(self, '警告', "覚えられていない単語はありません、他の選択肢を選んで下さい", QMessageBox.Ok)
      return

    #self.All_Erase2()

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1
    Question_Number = number
    Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    #self.All_Erase3()

    self.master.setCurrentIndex(3)
    #self.ShowResult(Memorized_count_in10)

    self.SaveResultMessage()

    self.Write10WordTest()

    self.UpdateFiles()

    TestWordList = [["" for column in range(4)] for row in range(10)]





  def MemorizedWordTest(self):
    global TestWordList
    global Question_Number
    global Memorized_count_in10

    Memorized_number = self.Extract10Words("NotMemorized")

    if Memorized_number ==0:
      msgBox0 = QMessageBox.warning(self, '警告', "覚えた単語はありません、他の選択肢を選んで下さい", QMessageBox.Ok)
      return

    #self.All_Erase2()

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1
    Question_Number = number
    Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    #self.All_Erase3()

    self.ShowResult(Memorized_count_in10)

    self.SaveResultMessage()

    self.Write10WordTest()

    self.UpdateFiles()

    TestWordList = [["" for column in range(4)] for row in range(10)]


  def GoBackToStartMenu(self):
    self.master.setCurrentIndex(0)


  def SaveResultMessage(self):
    msgBox = QMessageBox()
    msgBox.setText("結果をSaveしますか？")
    msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msgBox.setDefaultButton(QMessageBox.Yes)
    screenGeometry = QApplication.desktop().availableGeometry()
    screenGeo = screenGeometry.bottomRight()
    msgGeo = msgBox.frameGeometry()
    msgGeo.moveBottomRight(screenGeo)
    msgBox.show()
    ret = msgBox.exec_()


  def Write10WordTest(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Memorized_count_in10

    now = datetime.datetime.today()
    todaydatetime = now.strftime("%Y%m%d_%H:%M")

    #過去のファイル名前　10EngWordTest1 - 5の番号を+1する
    #ファイル名は10EngWordTest1 - 5にする。　日付はファイル名に使わない
    #この場合、historyファイルは使わない

    number = 0
    if JapEngWordIdiom_mode == 'EngWord':
      filenamepattern = re.compile(r'(Last5EngWordTest)(\d)(.dat)')
      filename = 'Last5EngWordTest'
    elif JapEngWordIdiom_mode == 'EngIdiom':
      filenamepattern = re.compile(r'(Last5EngIdiomTest)(\d)(.dat)')
      filename = 'Last5EngIdiomTest'
    elif JapEngWordIdiom_mode == 'JapWord':
      filenamepattern = re.compile(r'(Last5JapWordTest)(\d)(.dat)')
      filename = 'Last5JapWordTest'

    for filenames in os.listdir('.'):
      mo = filenamepattern.search(filenames)
      if mo == None:
        #print('LastTestのファイルは存在しません')
        self.NewlyCreateLastTest5File(Memorized_count_in10, todaydatetime)
        #break
      else:
        temp_number = number
        number = mo.group(2)
        number = max(temp_number, int(number))  #max関数の使い方、正しい？
        #print(number)

    if number == 5:
      filename = filename+'5'
      print(filename)
      #os.unlink(filename)

    else:
      for i in range(number):
        j = number - i
        filename1 = filename + str(j)
        filename2 = filename + str(j+1)
        print(filename1, filename2)
        #shutil.move(filename1, filename2)

    filename3 = filename + "1"
    shelf_file2 = shelve.open(filename3)
    shelf_file2['word'] = TestWordList
    shelf_file2['date'] = todaydatetime
    shelf_file2['accuracy_rate'] = Memorized_count_in10
    shelf_file2['mode'] = JapEngWordIdiom_mode
    shelf_file2.close()

    y = os.path.exists("./LastTest.dat")
    if y != False:
      print(y)
      #os.unlink('LastTest.bak')
      #os.unlink('LastTest.dat')
      #os.unlink('LastTest.dir')
    filename4 = 'LastTest'
    shelf_file3 = shelve.open(filename4)
    shelf_file3['word'] = TestWordList
    shelf_file3['date'] = todaydatetime
    shelf_file3['accuracy_rate'] = Memorized_count_in10
    shelf_file3['mode'] = JapEngWordIdiom_mode
    shelf_file3.close()


  def UpdateFiles(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Question_Number

    now = datetime.datetime.today()
    todaydate = now.strftime("%Y/%m/%d")

    if JapEngWordIdiom_mode == 'EngWord':
      shelf_file1 = shelve.open('EngJapWordFile')
    elif JapEngWordIdiom_mode == 'EngIdiom':
      shelf_file1 = shelve.open('EngJapIdiomFile')
    elif JapEngWordIdiom_mode == 'JapWord':
      shelf_file1 = shelve.open('JapEngWordFile')

    self.wordlist1 = shelf_file1['word']
    shelf_file1['date'] = todaydate

    #print(self.wordlist)
    #print(shelf_file['word'])

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1

    Question_Number = number
    print(Question_Number)
    for i in range(Question_Number):
      print(i)
      print(TestWordList[i][3])
      #self.wordlist1[TestWordList[i][3]][2] = 'Y'  #already_read
      #self.wordlist1[TestWordList[i][3]][3] += 1   #trial number
      #self.wordlist1[TestWordList[i][3]][4] = TestWordList[i][2]#覚えた？　正解？

    shelf_file1['word'] = self.wordlist1
    shelf_file1.close()


    shelf_file2 = shelve.open('MasterFileCopy')
    self.wordlist2 = shelf_file2['word']
    shelf_file2['date'] = todaydate
    temp_max_row = shelf_file2['max_row']

    if (JapEngWordIdiom_mode == 'EngWord') or (JapEngWordIdiom_mode == 'EngIdiom'):
      for j in range(Question_Number):
        for i in range(temp_max_row):
          if self.wordlist2[i][0] == TestWordList[j][0]:
            self.wordlist2[i][3]  = 'Y'  #already_read
            self.wordlist2[i][4]  += 1   #trial number
            self.wordlist2[i][5]  = TestWordList[j][2] #覚えた？　正解？
            continue
    elif JapEngWordIdiom_mode == 'JapWord':
      for j in range(Question_Number):
        for i in range(temp_max_row):
          if self.wordlist2[i][0] == TestWordList[j][1]:
            self.wordlist2[i][3]  = 'Y'  #already_read
            self.wordlist2[i][4]  += 1   #trial number
            self.wordlist2[i][5]  = TestWordList[j][2] #覚えた？　正解？
            continue
    shelf_file2['word'] = self.wordlist2
    shelf_file2.close()

  def NewlyCreateLastTest5File(self, accuracy, today):
    global TestWordList
    global JapEngWordIdiom_mode

    temp_accuracy = accuracy
    todaydatetime = today

    if JapEngWordIdiom_mode == 'EngWord':
      filename = "Last5EngWordTest1"
    elif JapEngWordIdiom_mode == 'EngIdiom':
      filename = "Last5EngIdiomTest1"
    elif JapEngWordIdiom_mode == 'JapWord':
      filename = "Last5JapWordTest1"

    shelf_file = shelve.open(filename)

    shelf_file['word'] = TestWordList
    shelf_file['date'] = todaydatetime
    shelf_file['accuracy_rate'] = temp_accuracy
    shelf_file['trial_num'] = 1

    #過去の同じ１０問ファイルを何回を行ったかを調べ
    #shelf_file1['trial_num']をUpdateする
    shelf_file.close()


  def UpdateFiles(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Question_Number

    now = datetime.datetime.today()
    todaydate = now.strftime("%Y/%m/%d")

    if JapEngWordIdiom_mode == 'EngWord':
      shelf_file1 = shelve.open('EngJapWordFile')
    elif JapEngWordIdiom_mode == 'EngIdiom':
      shelf_file1 = shelve.open('EngJapIdiomFile')
    elif JapEngWordIdiom_mode == 'JapWord':
      shelf_file1 = shelve.open('JapEngWordFile')

    self.wordlist1 = shelf_file1['word']
    shelf_file1['date'] = todaydate

    #print(self.wordlist)
    #print(shelf_file['word'])

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1

    Question_Number = number
    print(Question_Number)
    for i in range(Question_Number):
      print(i)
      print(TestWordList[i][3])
      #self.wordlist1[TestWordList[i][3]][2] = 'Y'  #already_read
      #self.wordlist1[TestWordList[i][3]][3] += 1   #trial number
      #self.wordlist1[TestWordList[i][3]][4] = TestWordList[i][2]#覚えた？　正解？

    shelf_file1['word'] = self.wordlist1
    shelf_file1.close()


    shelf_file2 = shelve.open('MasterFileCopy')
    self.wordlist2 = shelf_file2['word']
    shelf_file2['date'] = todaydate
    temp_max_row = shelf_file2['max_row']

    if (JapEngWordIdiom_mode == 'EngWord') or (JapEngWordIdiom_mode == 'EngIdiom'):
      for j in range(Question_Number):
        for i in range(temp_max_row):
          if self.wordlist2[i][0] == TestWordList[j][0]:
            self.wordlist2[i][3]  = 'Y'  #already_read
            self.wordlist2[i][4]  += 1   #trial number
            self.wordlist2[i][5]  = TestWordList[j][2] #覚えた？　正解？
            continue
    elif JapEngWordIdiom_mode == 'JapWord':
      for j in range(Question_Number):
        for i in range(temp_max_row):
          if self.wordlist2[i][0] == TestWordList[j][1]:
            self.wordlist2[i][3]  = 'Y'  #already_read
            self.wordlist2[i][4]  += 1   #trial number
            self.wordlist2[i][5]  = TestWordList[j][2] #覚えた？　正解？
            continue
    shelf_file2['word'] = self.wordlist2
    shelf_file2.close()





class RecentFileListup(QWidget):
  def __init__(self, parent):
    super().__init__(parent)
    self.master = parent

    self.label = QLabel(self)
    pixmap = QPixmap('AtFacebook.jpg')
    self.label.setPixmap(pixmap)
    self.label.move(410, 40)

    self.label2 = QLabel(self)
    self.label2.setText("Copyright 2018 Hiromi Maeda")
    self.label2.move(900,840)

    self.label9 = QLabel(self)
    self.label9.setText("番号　   　前回のテスト日時　      　　単語　　　　                                                                                       　前回の正解数 ")
    self.label9.move(50, 200)

    #self.label10 = QLabel(self)
    #self.label10.setText("    ")
    #self.label10.move(50, 200)

    #self.label11 = QLabel(self)
    #self.label11.setText(" ")
    #self.label11.move(50, 260)


    self.textbox11 = QTextEdit(self)
    self.textbox11.move(50, 250)
    self.textbox11.resize(30,30)

    self.textbox12 = QTextEdit(self)
    self.textbox12.move(100, 250)
    self.textbox12.resize(120,30)

    self.textbox13 = QTextEdit(self)
    self.textbox13.move(240, 250)
    self.textbox13.resize(400,30)

    self.textbox14 = QTextEdit(self)
    self.textbox14.move(680, 250)
    self.textbox14.resize(50,30)

    self.textbox21 = QTextEdit(self)
    self.textbox21.move(50, 300)
    self.textbox21.resize(30,30)

    self.textbox22 = QTextEdit(self)
    self.textbox22.move(100, 300)
    self.textbox22.resize(120,30)


    self.textbox23 = QTextEdit(self)
    self.textbox23.move(240, 300)
    self.textbox23.resize(400,30)

    self.textbox24 = QTextEdit(self)
    self.textbox24.move(680, 300)
    self.textbox24.resize(50,30)

    self.textbox31 = QTextEdit(self)
    self.textbox31.move(50, 350)
    self.textbox31.resize(30,30)

    self.textbox32 = QTextEdit(self)
    self.textbox32.move(100, 350)
    self.textbox32.resize(120,30)

    self.textbox33 = QTextEdit(self)
    self.textbox33.move(240, 350)
    self.textbox33.resize(400,30)

    self.textbox34 = QTextEdit(self)
    self.textbox34.move(680, 350)
    self.textbox34.resize(50,30)

    self.textbox41 = QTextEdit(self)
    self.textbox41.move(50, 400)
    self.textbox41.resize(30,30)

    self.textbox42 = QTextEdit(self)
    self.textbox42.move(100, 400)
    self.textbox42.resize(120,30)

    self.textbox43 = QTextEdit(self)
    self.textbox43.move(240, 400)
    self.textbox43.resize(400,30)

    self.textbox44 = QTextEdit(self)
    self.textbox44.move(680, 400)
    self.textbox44.resize(50,30)

    self.textbox51 = QTextEdit(self)
    self.textbox51.move(50, 450)
    self.textbox51.resize(30,30)

    self.textbox52 = QTextEdit(self)
    self.textbox52.move(100, 450)
    self.textbox52.resize(120,30)

    self.textbox53 = QTextEdit(self)
    self.textbox53.move(240, 450)
    self.textbox53.resize(400,30)

    self.textbox54 = QTextEdit(self)
    self.textbox54.move(680, 450)
    self.textbox54.resize(50,30)

    self.textbox61 = QTextEdit(self)
    self.textbox61.move(50, 500)
    self.textbox61.resize(30,30)

    self.textbox62 = QTextEdit(self)
    self.textbox62.move(100, 500)
    self.textbox62.resize(300,30)



  def SelectRecentFile(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Number_RecentTestFile
    global ExistingRecentTestFilename
    global Question_Number

    if JapEngWordIdiom_mode == 'EngWord':
      filenamepattern = re.compile(r'(Last5EngWordTest)(\d)(.dat)')
      filename = 'Last5EngWordTest'
    elif JapEngWordIdiom_mode == 'EngIdiom':
      filenamepattern = re.compile(r'(Last5EngIdiomTest)(\d)(.dat)')
      filename = 'Last5EngIdiomTest'
    elif JapEngWordIdiom_mode == 'JapWord':
      filenamepattern = re.compile(r'(Last5JapWordTest)(\d)(.dat)')
      filename = 'Last5JapWordTest'

    eachfilename = ["" for column in range(5)]
    number = 0
    i = 0
    for filenames in os.listdir('.'):
      mo = filenamepattern.search(filenames)
      if mo != None:
        number = mo.group(2)
        ExistingRecentTestFilename[i]  = mo.group(1)+ mo.group(2)
        shelf_file = shelve.open(ExistingRecentTestFilename[i])
        TestWordList=shelf_file['word']
        todaydatetime = shelf_file['date']
        accuracy = shelf_file['accuracy_rate']

        shelf_file.close()

        Question_Number = 0
        joined_word = TestWordList[0][0]
        for j in range(9):
          if TestWordList[j][0] == "":
            break
          else:
            joined_word = joined_word+", "+ TestWordList[j+1][0]  #文字列をつなぐ
            Question_Number += 1

        if i == 0:
          self.textbox11.setText('1')
          self.textbox11.show()
          self.textbox12.setText(todaydatetime)
          self.textbox12.show()
          self.textbox13.setText(joined_word)
          self.textbox13.show()
          self.textbox14.setText(str(accuracy))
          self.textbox14.show()
        elif i == 1:
          self.textbox21.setText('2')
          self.textbox21.show()
          self.textbox22.setText(todaydatetime)
          self.textbox22.show()
          self.textbox23.setText(joined_word)
          self.textbox23.show()
          self.textbox24.setText(str(accuracy))
          self.textbox24.show()
        elif i == 2:
          self.textbox31.setText('3')
          self.textbox31.show()
          self.textbox32.setText(todaydatetime)
          self.textbox32.show()
          self.textbox33.setText(joined_word)
          self.textbox33.show()
          self.textbox34.setText(str(accuracy))
          self.textbox34.show()
        elif i == 3:
          self.textbox41.setText('4')
          self.textbox41.show()
          self.textbox42.setText(todaydatetime)
          self.textbox42.show()
          self.textbox43.setText(joined_word)
          self.textbox43.show()
          self.textbox44.setText(str(accuracy))
          self.textbox44.show()
        elif i == 4:
          self.textbox51.setText('5')
          self.textbox51.show()
          self.textbox52.setText(todaydatetime)
          self.textbox52.show()
          self.textbox53.setText(joined_word)
          self.textbox53.show()
          self.textbox54.setText(str(accuracy))
          self.textbox54.show()
        i += 1
    self.textbox61.setText('6')
    self.textbox61.show()
    self.textbox62.setText("中断します")
    self.textbox62.show()
    Number_RecentTestFile = i

    #Dialog表示　　1-5を入力させる
    self.le = QLineEdit(self)
    self.le.move(650, 450)
    self.le.show

    self.button10.show()
    self.button10.clicked.connect(self.NumberInput_Selct_ConductTest)
                             #SelectedRecentFilenameIndexに何番が選ばれたかを代入



  def TestSelectedRecentFile(self):
    global TestWordList
    global SelectedRecentFilenameIndex
    global ExistingRecentTestFilename
    global Question_Number

    #self.All_Erase4()
    #self.setGeometry(610, 600, 200, 150)
    #self.setWindowTitle('番号入力画面')
    #self.show()

    #self.All_Erase3()


    filename = ExistingRecentTestFilename[SelectedRecentFilenameIndex]
    shelf_file = shelve.open(filename)
    TestWordList = shelf_file['word']
    print(TestWordList)
    todaydatetime = shelf_file['date']
    count = shelf_file['accuracy_rate']
    shelf_file.close()

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1
    Question_Number = number
    #これ以降は確認要する
    #RecentTest()の場合、Trial_numをUpdateする必要有
    print("Question_Numberは"+str(number))

    Memorized_count_in10 = self.ShowOneByOne()

    if Memorized_count_in10 == -1:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      #self.All_Erase3()
      return

    #self.All_Erase4()

    self.master.setCurrentIndex(3)
    #self.ShowResult(Memorized_count_in10)

    self.SaveResultMessage()

    self.RewriteForRecentTest()

    self.UpdateFiles()

    TestWordList = [["" for column in range(4)] for row in range(10)]






  def NumberInput_Selct_ConductTest(self):
    global Number_RecentTestFile
    global SelectedRecentFilenameIndex

    #self.All_Erase4()
    number, ok = QInputDialog.getText(self, 'Input Dialog',
            '選択したい番号を入力下さい:')
    if ok:
      self.le.setText(number)
      print(number)

    flag = True
    while(flag == True):
      if ((int(number) >= 1) and (int(number) <= Number_RecentTestFile)) or (int(number) == 6):
        break
      else:
        msgBox = QMessageBox.warning(self, '警告', "選択された番号は正しくありません、再度入力下さい", QMessageBox.Ok)

    SelectedRecentFilenameIndex = number



class QuestionAnswer(QWidget):
    #"""
    #役割 : Buttonクリック・イベントを設定 --> 質問と回答を表示する関数に飛ぶ
    #引数 : 無
    #戻り値  : 無
    #"""
  def __init__(self, parent):
    super().__init__(parent)
    self.master = parent

    self.label = QLabel(self)
    pixmap = QPixmap('AtFacebook.jpg')
    self.label.setPixmap(pixmap)
    self.label.move(410, 40)


    self.label2 = QLabel(self)
    self.label2.setText("Copyright 2018 Hiromi Maeda")
    self.label2.move(900,840)


    self.label21 = QLabel(self)
    self.label21.setText("それでは始めます。　下記ボタンを押して下さい！")
    self.label21.move(380,220)
    self.label21.show()

    self.button21 = QPushButton('開始！', self)
    self.button21.move (450, 250)
    self.button21.show()
    self.button21.clicked.connect(self.ShowOneByOne)


  def ShowOneByOne(self): #10問の内Memorizeした単語数をreturnする
    """
    役割 : 質問を１つ１つ出していく、都度「回答表示しますか？」表示、回答表示、「答が正しかったですか？」表示、MemorizedフラグON、質問の最後まで到達すると結果表示Window（Class）に切り替え
    引数 : 無
    戻り値  : 無
    """
    global TestWordList
    global Question_Number
    global Memorized_count_in10
    global ShowOneByOnePassTime

    self.label21.hide()
    self.button21.hide()

    print("entering ShowOneByOne")
    if ShowOneByOnePassTime == 0:
      self.label8 = QLabel(self)
      self.label8.move(380,320)
    self.label8.setText("１番目の問題です！　　　")

    #self.label8.show()

    self.textbox1 = QTextEdit(self)
    self.textbox1.move(380, 340)
    self.textbox1.resize(250,70)
    self.textbox1.show()

    if ShowOneByOnePassTime == 0:
      self.label10 = QLabel(self)
      self.label10.move(380,430)
    self.label10.setText("回答")

    self.label10.show()

    self.textbox2 = QTextEdit(self)
    self.textbox2.move(380, 450)
    self.textbox2.resize(250,70)
    self.textbox2.show()

    ShowOneByOnePassTime += 1

    Question_Number = min(10, Question_Number)

    Memorized_count_in10 = 0
    for i in range(Question_Number):
      self.label8.setText("               　　   　")
      self.label8.hide()
      self.label8.setText(str(i+1)+"番目の問題です！")
      self.label8.show()

      self.textbox1.setText(TestWordList[i][0])
      self.textbox1.show()
      self.textbox2.setText('        ')
      self.textbox2.show()

      msgBox1 = QMessageBox()
      msgBox1.setText("答を表示しますか？")
      msgBox1.setStandardButtons(QMessageBox.Ok)
      msgBox1.setDefaultButton(QMessageBox.Ok)
      screenGeometry1 = QApplication.desktop().availableGeometry()
      screenGeo1 = screenGeometry1.bottomRight()
      msgGeo1 = msgBox1.frameGeometry()
      msgGeo1.moveBottomRight(screenGeo1)
      msgBox1.show()
      ret1 = msgBox1.exec_()

      self.textbox2.setText(TestWordList[i][1])
      self.textbox2.show()

      msgBox2 = QMessageBox()
      msgBox2.setText("答は正しかったですか（”「覚えた」単語リスト”に加えますか）？")
      msgBox2.setStandardButtons(QMessageBox.Yes | QMessageBox.No | QMessageBox.Abort)
      msgBox2.setDefaultButton(QMessageBox.No)
      screenGeometry2 = QApplication.desktop().availableGeometry()
      screenGeo2 = screenGeometry2.bottomRight()
      msgGeo2 = msgBox2.frameGeometry()
      msgGeo2.moveBottomRight(screenGeo2)
      msgBox2.show()
      ret2 = msgBox2.exec_()

      #cancelの場合は、NotCompleteTestファイルを作る

      #self.textbox2.setText("")

      if ret2 == QMessageBox.Yes:
          Memorized_count_in10 += 1
          TestWordList[i][2] = 'Y'
      elif ret2 == QMessageBox.No:
          TestWordList[i][2] = 'N'
      elif ret2 == QMessageBox.Abort:
           Memorized_count_in10 = -1
           return

      if i == Question_Number-1:
        self.label8.setText("               　　 　")
        self.label8.setText("１番目の問題です！　　　")
        self.textbox1.setText(" ")
        self.textbox2.setText(" ")
        #self.label8.show()
        #self.textbox1.hide()
        #self.label10.hide()
        #self.textbox2.hide()

        self.master.setCurrentIndex(3)
        self.label21.show()  #QuestionAnswerの開始メッセージ
        self.button21.show()  #QuestionAnswerの開始ボタン

  #def DisplayAnswer(self, index):
    #temp_index = index
    #self.textbox2.setText(TestWordList[temp_index][1])




class ResultTable(QWidget):
    #"""
    #役割 : Buttonクリック・イベントを設置　-->１０問の質問、回答、正解・非正解の結果を表示する関数へ飛ぶ
    #引数 : 無
    #戻り値  : 無
    #"""
  def __init__(self, parent):
    super().__init__(parent)
    self.master = parent

    #self.label = QLabel(self)
    #pixmap = QPixmap('AtFacebook.jpg')
    #self.label.setPixmap(pixmap)
    #self.label.move(410, 40)
    #self.label.show()

    self.label2 = QLabel(self)
    self.label2.setText("Copyright 2018 Hiromi Maeda")
    self.label2.move(900,840)
    self.label2.show()

    self.label31 = QLabel(self)
    self.label31.setText("結果を表示します。　下記ボタンを押して下さい！")
    #self.label31.move(380,220)
    self.label31.move(380,700)

    self.button31 = QPushButton('結果表示！', self)
    #self.button31.move (450, 250)
    self.button31.move (450, 750)
    self.button31.clicked.connect(self.ShowResult)


  def ShowResult(self):
    """
    役割 : １０問の質問、回答、正解・非正解の結果を表示（SummaryData(), ShowTableData()関数を呼ぶ）、正解率計算・表示
    　　　　「結果をSaveするか？」表示、   Saveする為の関数（Write10WordTest()、UpdateFiles()）を呼び出す、TestWordListリセット（"")
           メインメニューへ戻る
    引数 :  無
    戻り値  無
    """
    global Memorized_count_in10
    global Question_Number
    global TestWordList
    global ShowOneByOnePassTime

    #if ShowOneByOnePassTime == 1:
      #self.label31.hide() #結果（表）を表示しますとのメッセージ
      #self.button31.hide() #結果表示ボタン
    #else:
      #print("Entering ResultDisplayMessageMode")
      #self.label31.show()
      #self.button31.show()

    self.SummaryData()
    self.ShowTableData()

    #self.label31.hide()
    #self.button31.hide()

    self.label32 = QLabel(self)
    self.label32.setText("あなたの正解率は")
    self.label32.move(330, 600)
    self.label32.show()

    self.label33 = QLabel(self)
    self.label33.setText(" ")
    self.label33.move(430, 600)
    self.label33.setText(str(Memorized_count_in10)+' / ' + str(Question_Number))
    self.label33.show()

    SaveOrNot = self.SaveResultMessage()

    if SaveOrNot == "Save":
      self.Write10WordTest()
      self.UpdateFiles()
      TestWordList = [["" for column in range(4)] for row in range(10)]
      self.SummaryData()
      self.ShowTableData()
      self.master.setCurrentIndex(0)
    else:
      TestWordList = [["" for column in range(4)] for row in range(10)]
      self.SummaryData()
      self.ShowTableData()
      self.master.setCurrentIndex(0)


  def SummaryData(self):
    """
    役割 : TestWordListの１０問の結果を、表に代入、表を表示する関数（ShowTableData())に飛ぶ
    引数 :  無
    戻り値  無
    """
    global TestWordList
    global Question_Number
    global ShowOneByOnePassTime

    self.model = QStandardItemModel(10, 3)

    row = 0
    col = 0

    for row in range(Question_Number):
      for col in range(3):
        item = QStandardItem(TestWordList[row][col])
        self.model.setItem( row, col, item)
    if ShowOneByOnePassTime >= 2:
      self.tv.hide()


  def ShowTableData(self):
    global JapEngWordIdiom_mode
    global ShowOneByOnePassTime

    #if ShowOneByOnePassTime == 1:
    if ShowOneByOnePassTime >= 0:
      self.tv = QTableView(self)
    self.tv.setModel(self.model)

    if (JapEngWordIdiom_mode == 'EngWord') or (JapEngWordIdiom_mode == 'EngIdiom'):
      self.tv.setColumnWidth(0, 200)
      self.tv.setColumnWidth(1, 600)
    else:
      self.tv.setColumnWidth(0, 600)
      self.tv.setColumnWidth(1, 200)

    self.tv.setGeometry(0, 0, 1000, 450)
    self.tv.show()


  def SaveResultMessage(self):
    msgBox = QMessageBox()
    msgBox.setText("結果をSaveしますか？")
    msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msgBox.setDefaultButton(QMessageBox.Yes)
    screenGeometry = QApplication.desktop().availableGeometry()
    screenGeo = screenGeometry.bottomRight()
    msgGeo = msgBox.frameGeometry()
    msgGeo.moveBottomRight(screenGeo)
    msgBox.show()
    ret = msgBox.exec_()

    if ret == QMessageBox.Yes:
          SaveOrNot = "Save"
    elif ret == QMessageBox.No:
          SaveOrNot = "NotSave"

    return SaveOrNot



  def UpdateFiles(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Question_Number

    now = datetime.datetime.today()
    todaydate = now.strftime("%Y/%m/%d")

    if JapEngWordIdiom_mode == 'EngWord':
      shelf_file1 = shelve.open('EngJapWordFile')
    elif JapEngWordIdiom_mode == 'EngIdiom':
      shelf_file1 = shelve.open('EngJapIdiomFile')
    elif JapEngWordIdiom_mode == 'JapWord':
      shelf_file1 = shelve.open('JapEngWordFile')

    self.wordlist1 = shelf_file1['word']
    shelf_file1['date'] = todaydate

    #print(self.wordlist)
    #print(shelf_file['word'])

    number = 0
    for i in range(10):
      if TestWordList[i][0] == "":
        break
      else:
        number += 1

    Question_Number = number
    print(Question_Number)
    for i in range(Question_Number):
      print(i)
      print(TestWordList[i][3])
      #self.wordlist1[TestWordList[i][3]][2] = 'Y'  #already_read
      #self.wordlist1[TestWordList[i][3]][3] += 1   #trial number
      #self.wordlist1[TestWordList[i][3]][4] = TestWordList[i][2]#覚えた？　正解？

    shelf_file1['word'] = self.wordlist1
    shelf_file1.close()


    shelf_file2 = shelve.open('MasterFileCopy')
    self.wordlist2 = shelf_file2['word']
    shelf_file2['date'] = todaydate
    temp_max_row = shelf_file2['max_row']

    if (JapEngWordIdiom_mode == 'EngWord') or (JapEngWordIdiom_mode == 'EngIdiom'):
      for j in range(Question_Number):
        for i in range(temp_max_row):
          if self.wordlist2[i][0] == TestWordList[j][0]:
            self.wordlist2[i][3]  = 'Y'  #already_read
            self.wordlist2[i][4]  += 1   #trial number
            self.wordlist2[i][5]  = TestWordList[j][2] #覚えた？　正解？
            continue
    elif JapEngWordIdiom_mode == 'JapWord':
      for j in range(Question_Number):
        for i in range(temp_max_row):
          if self.wordlist2[i][0] == TestWordList[j][1]:
            self.wordlist2[i][3]  = 'Y'  #already_read
            self.wordlist2[i][4]  += 1   #trial number
            self.wordlist2[i][5]  = TestWordList[j][2] #覚えた？　正解？
            continue
    shelf_file2['word'] = self.wordlist2
    shelf_file2.close()



  def Write10WordTest(self):
    """
    役割 :
    引数 :  無
    戻り値  無
    """
    global TestWordList
    global JapEngWordIdiom_mode
    global Memorized_count_in10

    now = datetime.datetime.today()
    todaydatetime = now.strftime("%Y%m%d_%H:%M")

    #過去のファイル名前　10EngWordTest1 - 5の番号を+1する
    #ファイル名は10EngWordTest1 - 5にする。　日付はファイル名に使わない
    #この場合、historyファイルは使わない

    number = 0
    if JapEngWordIdiom_mode == 'EngWord':
      filenamepattern = re.compile(r'(Last5EngWordTest)(\d)(.dat)')
      filename = 'Last5EngWordTest'
    elif JapEngWordIdiom_mode == 'EngIdiom':
      filenamepattern = re.compile(r'(Last5EngIdiomTest)(\d)(.dat)')
      filename = 'Last5EngIdiomTest'
    elif JapEngWordIdiom_mode == 'JapWord':
      filenamepattern = re.compile(r'(Last5JapWordTest)(\d)(.dat)')
      filename = 'Last5JapWordTest'

    for filenames in os.listdir('.'):
      mo = filenamepattern.search(filenames)
      if mo == None:
        #print('LastTestのファイルは存在しません')
        self.NewlyCreateLastTest5File(Memorized_count_in10, todaydatetime)
        #break
      else:
        temp_number = number
        number = mo.group(2)
        number = max(temp_number, int(number))  #max関数の使い方、正しい？
        #print(number)

    if number == 5:
      filename = filename+'5'
      print(filename)
      #os.unlink(filename)

    else:
      for i in range(number):
        j = number - i
        filename1 = filename + str(j)
        filename2 = filename + str(j+1)
        print(filename1, filename2)
        #shutil.move(filename1, filename2)

    filename3 = filename + "1"
    shelf_file2 = shelve.open(filename3)
    shelf_file2['word'] = TestWordList
    shelf_file2['date'] = todaydatetime
    shelf_file2['accuracy_rate'] = Memorized_count_in10
    shelf_file2['mode'] = JapEngWordIdiom_mode
    shelf_file2.close()

    y = os.path.exists("./LastTest.dat")
    if y != False:
      print(y)
      #os.unlink('LastTest.bak')
      #os.unlink('LastTest.dat')
      #os.unlink('LastTest.dir')
    filename4 = 'LastTest'
    shelf_file3 = shelve.open(filename4)
    shelf_file3['word'] = TestWordList
    shelf_file3['date'] = todaydatetime
    shelf_file3['accuracy_rate'] = Memorized_count_in10
    shelf_file3['mode'] = JapEngWordIdiom_mode
    shelf_file3.close()


  def NewlyCreateLastTest5File(self, accuracy, today):
    global TestWordList
    global JapEngWordIdiom_mode

    temp_accuracy = accuracy
    todaydatetime = today

    if JapEngWordIdiom_mode == 'EngWord':
      filename = "Last5EngWordTest1"
    elif JapEngWordIdiom_mode == 'EngIdiom':
      filename = "Last5EngIdiomTest1"
    elif JapEngWordIdiom_mode == 'JapWord':
      filename = "Last5JapWordTest1"

    shelf_file = shelve.open(filename)

    shelf_file['word'] = TestWordList
    shelf_file['date'] = todaydatetime
    shelf_file['accuracy_rate'] = temp_accuracy
    shelf_file['trial_num'] = 1

    #過去の同じ１０問ファイルを何回を行ったかを調べ
    #shelf_file1['trial_num']をUpdateする
    shelf_file.close()



  def RewriteForRecentTest(self, accuracy):
    global TestWordList
    global JapEngWordIdiom_mode
    global SelectedRecentFilenameIndex
    temp_accuracy = accuracy

    now = datetime.datetime.today()
    todaydatetime = now.strftime("%Y%m%d_%H:%M")

    y = os.path.exists("./LastTest.dat")
    if y != False:
      print(y)
      #os.unlink('LastTest.bak')
      #os.unlink('LastTest.dat')
      #os.unlink('LastTest.dir')

    #LastTestファイルのUpdate
    shelf_file1 = shelve.open('LastTest')
    shelf_file1['word'] = TestWordList
    shelf_file1['date'] = todaydatetime
    shelf_file1['accuracy_rate'] = temp_accuracy
    shelf_file1['mode'] = JapEngWordIdiom_mode
    shelf_file1.close()

    number = 0
    if JapEngWordIdiom_mode == 'EngWord':
      filename0 = 'Last5EngWordTest'
    elif JapEngWordIdiom_mode == 'EngIdiom':
      filename0 = 'Last5EngIdiomTest'
    elif JapEngWordIdiom_mode == 'JapWord':
      filename0 = 'Last5JapWordTest'


    #選択されたファイルを、ファイル名の最後の数字1に移動、その分、他のファイルを+1shift
    for i in range(SelectedRecentFilenameIndex-1):
      filename1 = filename0 + str(SelectedRecentFilenameIndex-i)
      filename2 = filename0 + str(SelectedRecentFilenameIndex-i-1)
      print(filename2, filename1)
      #shutil.move(filename2, filename1)

    filename = filename0 + str(1)
    shelf_file2 = shelve.open(filename)
    shelf_file2['word'] = TestWordList
    shelf_file2['date'] = todaydatetime
    shelf_file2['accuracy_rate'] = temp_accuracy
    shelf_file2['trial_num'] += 1

    shelf_file2.close()




  def UpdateFiles_After_Redoing_Last10WordTest(self):
    global TestWordList
    global JapEngWordIdiom_mode
    global Question_Number


    now = datetime.datetime.today()
    todaydate = now.strftime("%Y/%m/%d")


    if JapEngWordIdiom_mode == 'EngWord':
      shelf_file1 = shelve.open('EngJapWordFile')
    elif JapEngWordIdiom_mode == 'EngIdiom':
      shelf_file1 = shelve.open('EngJapIdiomFile')
    elif JapEngWordIdiom_mode == 'JapWord':
      shelf_file1 = shelve.open('JapEngWordFile')

    self.wordlist1 = shelf_file1['word']
    shelf_file1['date'] = todaydate

    #print(self.wordlist)
    #print(shelf_file['word'])

    for i in range(Question_Number):
      self.wordlist1[TestWordList[i][3]][2] = 'Y'  #already_read
      self.wordlist1[TestWordList[i][3]][3] += 1   #trial number
      self.wordlist1[TestWordList[i][3]][4] = TestWordList[i][2]#覚えた？　正解？

    shelf_file1['word'] = self.wordlist1
    shelf_file1.close()


    shelf_file2 = shelve.open('MasterFileCopy')
    self.wordlist2 = shelf_file2['word']
    shelf_file2['date'] = todaydate
    temp_max_row = shelf_file2['max_row']

    if (JapEngWordIdiom_mode == 'EngWord') or (JapEngWordIdiom_mode == 'EngIdiom'):
      for i in range(temp_max_row):
        if self.wordlist2[i][0] == TestWordList[i][0]:
          self.wordlist2[i][3]  = 'Y'  #already_read
          self.wordlist2[i][4]  += 1   #trial number
          self.wordlist2[i][5]  = TestWordList[i][2] #覚えた？　正解？
    elif JapEngWordIdiom_mode == 'JapWord':
      for i in range(temp_max_row):
        if self.wordlist2[i][0] == TestWordList[i][1]:
          self.wordlist2[i][3]  = 'Y'  #already_read
          self.wordlist2[i][4]  += 1   #trial number
          self.wordlist2[i][5]  = TestWordList[i][2] #覚えた？　正解？

    shelf_file2['word'] = self.wordlist2
    shelf_file2.close()


class FirstFilePreparation:
  def __init__(self):
    global Original_Excel_Max_Row_Number

    now = datetime.datetime.today()
    self.todaydate = now.strftime("%Y/%m/%d")

    wb = openpyxl.load_workbook('EnglishWOrdIdiomMasterFile.xlsx')
    self.sheet = wb.get_sheet_by_name('Sheet1')

    Original_Excel_Max_Row_Number = self.sheet.max_row

    self.MasterWordlist = [["" for column in range(6)] for row in range(Original_Excel_Max_Row_Number+1)]
    #for row in range(max_row_number+1):
      #self.MasterWordlist[row][4] = 0
      #self.MasterWordlist[row][5] = ""

    y = os.path.exists("./MasterFileCopy.dat")
    if y == False:
      self.FirstMasterFileCopy()
      self.DivideMasterCreateIndividual()


  def FirstMasterFileCopy(self):

    tested = 'N'
    trial_num = 0
    memorized = '0'
    for row_num in range(1, self.max_row_number+1):
      self.MasterWordlist[row_num][0] = self.sheet.cell(row=row_num, column=1).value #Eng
      self.MasterWordlist[row_num][1] = self.sheet.cell(row=row_num, column=2).value #Jap
      self.MasterWordlist[row_num][2] = self.sheet.cell(row=row_num, column=3).value #Idiom?
      self.MasterWordlist[row_num][3] = tested #tested
      self.MasterWordlist[row_num][4] = trial_num  #number of trial
      self.MasterWordlist[row_num][5] = memorized #memorized

    shelf_file = shelve.open('MasterFileCopy')
    shelf_file['word'] = self.MasterWordlist
    shelf_file['max_row'] = self.max_row_number
    shelf_file['date'] = self.todaydate
    shelf_file.close()


  def DivideMasterCreateIndividual(self):
    shelf_file0 = shelve.open('MasterFileCopy')
    self.MasterWordlist = shelf_file0['word']

    counter0 = 0
    for row_num in range(1, self.max_row_number+1):
      if self.MasterWordlist[row_num][2] == "I":
        counter0 += 1
        print(counter0)

    shelf_file1 = shelve.open('EngJapWordFile')
    shelf_file2 = shelve.open('EngJapIdiomFile')
    shelf_file3 = shelve.open('JapEngWordFile')

    self.wordlistEngWord = [["" for column in range(5)] for row in range(self.max_row_number-counter0)]
    self.wordlistEngIdiom = [["" for column in range(5)] for row in range(counter0)]
    self.wordlistJapWord = [["" for column in range(5)] for row in range(self.max_row_number+1)]

    counter1 = 0
    counter2 = 0
    for row_num in range(1, self.max_row_number+1):
      if self.MasterWordlist[row_num][2] == "I":
        self.wordlistEngIdiom[counter2][0] = self.MasterWordlist[row_num][0]
        self.wordlistEngIdiom[counter2][1] = self.MasterWordlist[row_num][1]
        self.wordlistEngIdiom[counter2][2] = self.MasterWordlist[row_num][3]
        self.wordlistEngIdiom[counter2][3] = self.MasterWordlist[row_num][4]
        self.wordlistEngIdiom[counter2][4] = self.MasterWordlist[row_num][5]
        counter2 += 1
      else:
        self.wordlistEngWord[counter1][0] = self.MasterWordlist[row_num][0]
        self.wordlistEngWord[counter1][1] = self.MasterWordlist[row_num][1]
        self.wordlistEngWord[counter1][2] = self.MasterWordlist[row_num][3]
        self.wordlistEngWord[counter1][3] = self.MasterWordlist[row_num][4]
        self.wordlistEngWord[counter1][4] = self.MasterWordlist[row_num][5]
        counter1 += 1

      self.wordlistJapWord[row_num][0] = self.MasterWordlist[row_num][1]
      self.wordlistJapWord[row_num][1] = self.MasterWordlist[row_num][0]
      self.wordlistJapWord[row_num][2] = self.MasterWordlist[row_num][3]
      self.wordlistJapWord[row_num][3] = self.MasterWordlist[row_num][4]
      self.wordlistJapWord[row_num][4] = self.MasterWordlist[row_num][5]

    shelf_file1['word'] = self.wordlistEngWord
    shelf_file1['max_row'] = counter1
    shelf_file1['date'] = self.todaydate
    shelf_file1.close()

    shelf_file2['word'] = self.wordlistEngIdiom
    shelf_file2['max_row'] = counter2
    shelf_file2['date'] = self.todaydate
    shelf_file2.close()

    shelf_file3['word'] = self.wordlistJapWord
    shelf_file3['max_row'] = self.max_row_number
    shelf_file3['date'] = self.todaydate
    shelf_file3.close()

class App(QTabWidget):
  def __init__(self):
    super().__init__()

    self.title = "Henryの英単語・イディオム Building BootCamp"
    self.left = 0
    self.top = 0
    self.width = 1100
    self.height = 900

    self.setWindowTitle(self.title)

    #１個1個のTabがメニュー画面に相当
    self.tab1 = StartMenu(self)
    self.tab2 = NextMenu(self)
    self.tab3 = QuestionAnswer(self)
    self.tab4 = ResultTable(self)
    self.tab5 = RecentFileListup(self)

    #タブページに追加
    self.addTab(self.tab1, "Startmenu")
    self.addTab(self.tab2, "Nextmenu")
    self.addTab(self.tab3, "QuestionAnswer")
    self.addTab(self.tab4, "ResultTable")
    self.addTab(self.tab5, "RecentFileListup")

    #TabパネルのBorderを削除
    self.setStyleSheet("QTabWidget::pane { border: 0; }")

    #self.tabBar().hide()
    self.resize(self.width, self.height)
    self.move(100, 0)



if __name__ == '__main__':

  #もしmydata.pyファイルが無ければというコマンドが必要
  FilePreparation = FirstFilePreparation()


  app = QApplication(sys.argv)
  ex1 = App()
  ex1.show()

  sys.exit(app.exec_())
