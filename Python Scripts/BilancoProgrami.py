# -*- coding: utf-8 -*-
"""
Created on Sat Feb  1 18:29:16 2020

@author: POI-PC
"""

from PyQt5.QtWidgets import*
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtGui
from PyQt5 import QtCore, QtWidgets
import sys
from selenium import webdriver
import time
import pandas as pd
import numpy as np
from xlrd import open_workbook
import os
from  openpyxl import *
import io
from zipfile import ZipFile
import xlrd
import codecs
import shutil
from selenium.common.exceptions import NoSuchElementException
import html5lib
from os import path
from pathlib import Path
from itertools import product
import xlwings as xw
from datetime import date



class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1489, 901)
        font = QtGui.QFont()
        font.setPointSize(9)
        MainWindow.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("icons/bilanco.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setLocale(QtCore.QLocale(QtCore.QLocale.Turkish, QtCore.QLocale.Turkey))
        MainWindow.setIconSize(QtCore.QSize(50, 50))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(380, 10, 711, 51))
        font = QtGui.QFont()
        font.setFamily("Tw Cen MT")
        font.setPointSize(24)
        font.setBold(True)
        font.setItalic(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.sirketGetir = QtWidgets.QPushButton(self.centralwidget)
        self.sirketGetir.setGeometry(QtCore.QRect(760, 120, 241, 61))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.sirketGetir.setFont(font)
        self.sirketGetir.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.sirketGetir.setLocale(QtCore.QLocale(QtCore.QLocale.Turkish, QtCore.QLocale.Turkey))
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("icons/sirketler.jpg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.sirketGetir.setIcon(icon1)
        self.sirketGetir.setIconSize(QtCore.QSize(50, 50))
        self.sirketGetir.setObjectName("sirketGetir")
        self.yedekleSil = QtWidgets.QPushButton(self.centralwidget)
        self.yedekleSil.setGeometry(QtCore.QRect(50, 120, 241, 61))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.yedekleSil.setFont(font)
        self.yedekleSil.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.yedekleSil.setLocale(QtCore.QLocale(QtCore.QLocale.Turkish, QtCore.QLocale.Turkey))
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("icons/clear.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.yedekleSil.setIcon(icon2)
        self.yedekleSil.setIconSize(QtCore.QSize(50, 50))
        self.yedekleSil.setObjectName("yedekleSil")
        self.anaExcel = QtWidgets.QPushButton(self.centralwidget)
        self.anaExcel.setGeometry(QtCore.QRect(1080, 120, 251, 61))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.anaExcel.setFont(font)
        self.anaExcel.setLocale(QtCore.QLocale(QtCore.QLocale.Turkish, QtCore.QLocale.Turkey))
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("icons/excel2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.anaExcel.setIcon(icon3)
        self.anaExcel.setIconSize(QtCore.QSize(50, 50))
        self.anaExcel.setObjectName("anaExcel")
        self.sirketler = QtWidgets.QListWidget(self.centralwidget)
        self.sirketler.setGeometry(QtCore.QRect(290, 290, 261, 301))
        self.sirketler.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.sirketler.setObjectName("sirketler")
        self.gosterSirket = QtWidgets.QPushButton(self.centralwidget)
        self.gosterSirket.setGeometry(QtCore.QRect(880, 330, 271, 61))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.gosterSirket.setFont(font)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("icons/show.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.gosterSirket.setIcon(icon4)
        self.gosterSirket.setIconSize(QtCore.QSize(40, 40))
        self.gosterSirket.setObjectName("gosterSirket")
        self.secilenIndr = QtWidgets.QPushButton(self.centralwidget)
        self.secilenIndr.setGeometry(QtCore.QRect(880, 490, 271, 61))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.secilenIndr.setFont(font)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap("icons/download.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.secilenIndr.setIcon(icon5)
        self.secilenIndr.setIconSize(QtCore.QSize(40, 40))
        self.secilenIndr.setObjectName("secilenIndr")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(60, 240, 191, 21))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.tumSirketler = QtWidgets.QListView(self.centralwidget)
        self.tumSirketler.setGeometry(QtCore.QRect(20, 290, 256, 301))
        self.tumSirketler.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.tumSirketler.setLineWidth(0)
        self.tumSirketler.setResizeMode(QtWidgets.QListView.Fixed)
        self.tumSirketler.setObjectName("tumSirketler")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(280, 240, 291, 16))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(290, 270, 261, 16))
        self.label_5.setObjectName("label_5")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(1150, 330, 20, 211))
        self.line.setLineWidth(3)
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setGeometry(QtCore.QRect(20, 200, 1381, 16))
        self.line_2.setLineWidth(2)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.bildirim = QtWidgets.QLineEdit(self.centralwidget)
        self.bildirim.setGeometry(QtCore.QRect(10, 800, 701, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.bildirim.setFont(font)
        self.bildirim.setObjectName("bildirim")
        self.genelGetir = QtWidgets.QPushButton(self.centralwidget)
        self.genelGetir.setGeometry(QtCore.QRect(1190, 330, 261, 61))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.genelGetir.setFont(font)
        icon6 = QtGui.QIcon()
        icon6.addPixmap(QtGui.QPixmap("icons/geneldown.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.genelGetir.setIcon(icon6)
        self.genelGetir.setIconSize(QtCore.QSize(35, 35))
        self.genelGetir.setObjectName("genelGetir")
        self.devamEt = QtWidgets.QPushButton(self.centralwidget)
        self.devamEt.setGeometry(QtCore.QRect(1190, 480, 261, 61))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.devamEt.setFont(font)
        icon13 = QtGui.QIcon()
        icon13.addPixmap(QtGui.QPixmap("icons/continue.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.devamEt.setIcon(icon13)
        self.devamEt.setIconSize(QtCore.QSize(35, 35))
        self.devamEt.setObjectName("devamEt")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(610, 70, 271, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setItalic(True)
        font.setUnderline(True)
        font.setWeight(75)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(980, 230, 331, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setItalic(True)
        font.setUnderline(True)
        font.setWeight(75)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(1050, 570, 251, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setItalic(True)
        font.setUnderline(True)
        font.setWeight(75)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.line_3 = QtWidgets.QFrame(self.centralwidget)
        self.line_3.setGeometry(QtCore.QRect(880, 550, 531, 20))
        self.line_3.setLineWidth(2)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.zipAktar = QtWidgets.QPushButton(self.centralwidget)
        self.zipAktar.setGeometry(QtCore.QRect(610, 660, 241, 61))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.zipAktar.setFont(font)
        icon7 = QtGui.QIcon()
        icon7.addPixmap(QtGui.QPixmap("icons/zip.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.zipAktar.setIcon(icon7)
        self.zipAktar.setIconSize(QtCore.QSize(40, 40))
        self.zipAktar.setObjectName("zipAktar")
        self.aktarExcel = QtWidgets.QPushButton(self.centralwidget)
        self.aktarExcel.setGeometry(QtCore.QRect(1160, 710, 241, 61))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.aktarExcel.setFont(font)
        icon8 = QtGui.QIcon()
        icon8.addPixmap(QtGui.QPixmap("icons/excel3.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.aktarExcel.setIcon(icon8)
        self.aktarExcel.setIconSize(QtCore.QSize(50, 50))
        self.aktarExcel.setObjectName("aktarExcel")
        self.label_12 = QtWidgets.QLabel(self.centralwidget)
        self.label_12.setGeometry(QtCore.QRect(1220, 640, 55, 16))
        font = QtGui.QFont()
        font.setPointSize(7)
        font.setBold(True)
        font.setWeight(75)
        self.label_12.setFont(font)
        self.label_12.setObjectName("label_12")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(1190, 610, 191, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setGeometry(QtCore.QRect(1300, 660, 51, 22))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label_13 = QtWidgets.QLabel(self.centralwidget)
        self.label_13.setGeometry(QtCore.QRect(1310, 640, 21, 16))
        font = QtGui.QFont()
        font.setPointSize(7)
        font.setBold(True)
        font.setWeight(75)
        self.label_13.setFont(font)
        self.label_13.setObjectName("label_13")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setGeometry(QtCore.QRect(1220, 660, 51, 22))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.bosZipler = QtWidgets.QListWidget(self.centralwidget)
        self.bosZipler.setGeometry(QtCore.QRect(350, 670, 241, 91))
        self.bosZipler.setObjectName("bosZipler")
        self.label_14 = QtWidgets.QLabel(self.centralwidget)
        self.label_14.setGeometry(QtCore.QRect(630, 740, 181, 21))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.label_14.setFont(font)
        self.label_14.setObjectName("label_14")
        self.secHepsini = QtWidgets.QPushButton(self.centralwidget)
        self.secHepsini.setGeometry(QtCore.QRect(420, 600, 131, 28))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.secHepsini.setFont(font)
        icon9 = QtGui.QIcon()
        icon9.addPixmap(QtGui.QPixmap("icons/selectall.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.secHepsini.setIcon(icon9)
        self.secHepsini.setObjectName("secHepsini")
        self.yedekle = QtWidgets.QPushButton(self.centralwidget)
        self.yedekle.setGeometry(QtCore.QRect(390, 120, 241, 61))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.yedekle.setFont(font)
        icon10 = QtGui.QIcon()
        icon10.addPixmap(QtGui.QPixmap("icons/backup.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.yedekle.setIcon(icon10)
        self.yedekle.setIconSize(QtCore.QSize(30, 30))
        self.yedekle.setObjectName("yedekle")
        self.listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget.setGeometry(QtCore.QRect(880, 640, 256, 192))
        self.listWidget.setObjectName("listWidget")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(1150, 790, 251, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setGeometry(QtCore.QRect(970, 290, 231, 21))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(1260, 290, 113, 22))
        self.lineEdit.setObjectName("lineEdit")
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setGeometry(QtCore.QRect(570, 260, 301, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")
        self.listWidget_2 = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget_2.setGeometry(QtCore.QRect(580, 290, 256, 301))
        self.listWidget_2.setObjectName("listWidget_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(880, 410, 271, 61))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        icon11 = QtGui.QIcon()
        icon11.addPixmap(QtGui.QPixmap("icons/checkbox.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton.setIcon(icon11)
        self.pushButton.setIconSize(QtCore.QSize(40, 40))
        self.pushButton.setObjectName("pushButton")
        self.label_15 = QtWidgets.QLabel(self.centralwidget)
        self.label_15.setGeometry(QtCore.QRect(10, 770, 171, 16))
        font = QtGui.QFont()
        font.setItalic(True)
        self.label_15.setFont(font)
        self.label_15.setObjectName("label_15")
        self.line_4 = QtWidgets.QFrame(self.centralwidget)
        self.line_4.setGeometry(QtCore.QRect(1470, 100, 20, 731))
        self.line_4.setLineWidth(3)
        self.line_4.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.label_16 = QtWidgets.QLabel(self.centralwidget)
        self.label_16.setGeometry(QtCore.QRect(940, 260, 261, 21))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_16.setFont(font)
        self.label_16.setObjectName("label_16")
        self.donem = QtWidgets.QLineEdit(self.centralwidget)
        self.donem.setGeometry(QtCore.QRect(1260, 260, 113, 22))
        self.donem.setObjectName("donem")
        self.tumGetir = QtWidgets.QPushButton(self.centralwidget)
        self.tumGetir.setGeometry(QtCore.QRect(1170, 400, 291, 61))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.tumGetir.setFont(font)
        icon12 = QtGui.QIcon()
        icon12.addPixmap(QtGui.QPixmap("icons/all.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.tumGetir.setIcon(icon12)
        self.tumGetir.setIconSize(QtCore.QSize(35, 35))
        self.tumGetir.setObjectName("tumGetir")
        self.label.raise_()
        self.sirketGetir.raise_()
        self.yedekleSil.raise_()
        self.anaExcel.raise_()
        self.sirketler.raise_()
        self.secilenIndr.raise_()
        self.label_3.raise_()
        self.tumSirketler.raise_()
        self.label_4.raise_()
        self.label_5.raise_()
        self.line.raise_()
        self.line_2.raise_()
        self.bildirim.raise_()
        self.gosterSirket.raise_()
        self.genelGetir.raise_()
        self.devamEt.raise_()
        self.label_7.raise_()
        self.label_8.raise_()
        self.label_9.raise_()
        self.line_3.raise_()
        self.zipAktar.raise_()
        self.aktarExcel.raise_()
        self.label_12.raise_()
        self.label_6.raise_()
        self.lineEdit_3.raise_()
        self.label_13.raise_()
        self.lineEdit_4.raise_()
        self.bosZipler.raise_()
        self.label_14.raise_()
        self.secHepsini.raise_()
        self.yedekle.raise_()
        self.listWidget.raise_()
        self.label_2.raise_()
        self.label_10.raise_()
        self.lineEdit.raise_()
        self.label_11.raise_()
        self.listWidget_2.raise_()
        self.pushButton.raise_()
        self.label_15.raise_()
        self.line_4.raise_()
        self.label_16.raise_()
        self.donem.raise_()
        self.tumGetir.raise_()
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1489, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.label_5.setBuddy(self.sirketler)
        self.retranslateUi(MainWindow)
        self.sirketGetir.clicked.connect(MainWindow.sirketlerKap)
        self.anaExcel.clicked.connect(MainWindow.bilancoExcel)
        self.yedekleSil.clicked.connect(MainWindow.silYedekle)
        self.gosterSirket.clicked.connect(self.sirketler.doItemsLayout)
        self.genelGetir.clicked.connect(MainWindow.genelYukle)
        self.anaExcel.released.connect(self.tumSirketler.doItemsLayout)
        self.devamEt.clicked.connect(MainWindow.devamEttir)
        self.zipAktar.clicked.connect(MainWindow.zipeAktar)
        self.aktarExcel.clicked.connect(MainWindow.hepsiExcel)
        self.zipAktar.released.connect(self.bosZipler.doItemsLayout)
        self.secilenIndr.clicked.connect(MainWindow.cekSecilen)
        self.secHepsini.clicked.connect(MainWindow.selectHepsi)
        self.secHepsini.clicked.connect(self.sirketler.selectAll)
        self.yedekle.clicked.connect(MainWindow.excelYedekle)
        self.aktarExcel.clicked.connect(self.listWidget.doItemsLayout)
        self.pushButton.clicked.connect(self.listWidget_2.doItemsLayout)
        self.tumGetir.clicked.connect(MainWindow.donemselTum)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Bilanco Programı"))
        self.label.setText(_translate("MainWindow", "Otomatik Bilanço Veri Çekme Programı V. 1.6"))
        self.sirketGetir.setText(_translate("MainWindow", "Şirketleri Getir"))
        self.yedekleSil.setText(_translate("MainWindow", "Sil ve Yedekle"))
        self.anaExcel.setText(_translate("MainWindow", "Bilanco.xlsx Sisteme Al"))
        self.gosterSirket.setText(_translate("MainWindow", "İndirilmemiş Verileri Göster"))
        self.secilenIndr.setText(_translate("MainWindow", "Seçileni İndir"))
        self.label_3.setText(_translate("MainWindow", "Tüm Şirketlerin Listesi"))
        self.label_4.setText(_translate("MainWindow", "Sistemde Çekilmemiş Şirketler Listesi"))
        self.label_5.setText(_translate("MainWindow", "(Burda tıkladıkların sisteme çekilecektir.)"))
        self.bildirim.setText(_translate("MainWindow", "Bildirimler !"))
        self.genelGetir.setText(_translate("MainWindow", "Tüm Şirketleri İndir"))
        self.devamEt.setText(_translate("MainWindow", "Kaldığı Yerden Devam Ettir"))
        self.label_7.setText(_translate("MainWindow", "Veriler İçin Ön Hazırlık"))
        self.label_8.setText(_translate("MainWindow", "Verilerin İnternetten Çekildiği Yer"))
        self.label_9.setText(_translate("MainWindow", "Verilerin Excel\'e Aktarılması"))
        self.zipAktar.setText(_translate("MainWindow", "Zip Dosyalarını Aç"))
        self.aktarExcel.setText(_translate("MainWindow", "Excel\'e Aktar"))
        self.label_12.setText(_translate("MainWindow", "Dönem"))
        self.label_6.setText(_translate("MainWindow", "Çekmek İstediğin Dönem"))
        self.lineEdit_3.setText(_translate("MainWindow", "2019"))
        self.label_13.setText(_translate("MainWindow", "Yıl"))
        self.lineEdit_4.setText(_translate("MainWindow", "0"))
        self.label_14.setText(_translate("MainWindow", "<-- Zip\'leri Boş Olanlar"))
        self.secHepsini.setText(_translate("MainWindow", "Hepsini Seç"))
        self.yedekle.setText(_translate("MainWindow", "Bilanco Yedekle"))
        self.label_2.setText(_translate("MainWindow", " <-- Excel\'e Aktarılmamış Olanlar"))
        self.label_10.setText(_translate("MainWindow", "İndirmek İstediğin Yılı Gir ->"))
        self.lineEdit.setText(_translate("MainWindow", "2020"))
        self.label_11.setText(_translate("MainWindow", "Seçilmiş (İndirilecek) Şirketler Listesi"))
        self.pushButton.setText(_translate("MainWindow", "Seçilmişleri Göster"))
        self.label_15.setText(_translate("MainWindow", "Writed by SVS © (2020)"))
        self.label_16.setText(_translate("MainWindow", "İndirmek İstediğin Dönemi Gir ->"))
        self.donem.setText(_translate("MainWindow", "5"))
        self.tumGetir.setText(_translate("MainWindow", "Tüm Şirketleri Dönemsel İndir"))




class Bilanco(QMainWindow):

    def __init__(self):
        super().__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.sirketler.setSelectionMode(
            QAbstractItemView.ExtendedSelection
        )
        self.ui.sirketler.setEditTriggers(QAbstractItemView.DoubleClicked|QAbstractItemView.EditKeyPressed)
        self.ui.sirketler.setSelectionMode(QAbstractItemView.MultiSelection)
        self.ui.sirketler.setViewMode(QListView.ListMode)
        self.ui.listWidget.setSelectionMode(
            QAbstractItemView.ExtendedSelection
        )
        self.ui.listWidget.setEditTriggers(QAbstractItemView.DoubleClicked|QAbstractItemView.EditKeyPressed)
        self.ui.listWidget.setSelectionMode(QAbstractItemView.MultiSelection)
        self.ui.listWidget.setViewMode(QListView.ListMode)
        self.ui.anaExcel.released.connect(self.listeyeDok)
        self.ui.sirketGetir.released.connect(self.bildirim1)
        #self.ui.anaExcel.released.connect(self.bildirim2)
        self.ui.yedekleSil.released.connect(self.bildirim3)
        self.ui.gosterSirket.clicked.connect(self.widgetListele)
        self.ui.gosterSirket.released.connect(self.bildirim4)
        self.ui.pushButton.clicked.connect(self.widgetSelectedShow)
        self.ui.pushButton.released.connect(self.bildirim8)
        #self.ui.sirketGetir.released.connect(self.listeyeDok)
        self.ui.zipAktar.released.connect(self.bildirim7)
        self.ui.sirketler.itemClicked.connect(self.seciliSec)
     
        
        
        
    def bildirim1(self):
        self.ui.bildirim.setText("Sirket Verileri Cekildi!")
    
    def bildirim2(self):
        self.ui.bildirim.setText("Excel Datası Cekildi!")
        
    def bildirim3(self):
        self.ui.bildirim.setText("Eski Veriler silindi ve Bilanco yedeklendi!")    
    
    def bildirim4(self):
        self.ui.bildirim.setText("Çekilen şirketler gösterildi!")
        
    def bildirim5(self):
        self.ui.bildirim.setText("Tum veriler CEKILEMEDI!")
    
    def bildirim6(self):    
        self.ui.bildirim.setText("Tum veriler basariyla cekildi!")
    
    def bildirim7(self):    
        self.ui.bildirim.setText("Dosyadaki tum Zip'ler açıldı!")
        
    def bildirim8(self):
        self.ui.bildirim.setText("Secilmis sirketler gösterildi!")
    
    def selectHepsi(self):
        print("ok")
        
    def excelYedekle(self):
        today = date.today()
        shutil.copy('Bilanco-Excel/Bilanco.xlsm', 'BilancoYedek/BilancoBackUp-'+str(today)+'.xlsm')
        self.ui.bildirim.setText("Bilanco excel'i yedeklendi!")
        
    def donemselTum(self):
        yil = int(self.ui.lineEdit.text())
        donem = int(self.ui.donem.text())
        
        yilDonem = str(yil) + "+" + str(donem)
        
        options = webdriver.ChromeOptions() 
        adres = fileName + "\Veriler\-"
        #options.add_argument("download.default_directory="+ adres ")
        prefs = {
        "download.default_directory": adres+yilDonem,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
        }
    
    
        options.add_experimental_option('prefs', prefs)
        browser = webdriver.Chrome(chrome_options=options)
        
        browser.get("https://www.kap.org.tr/tr/")
        time.sleep(5)
            
        ftablolar = browser.find_element_by_xpath("//*[@id='financialTablesTab']/div")
        ftablolar.click()
        
        time.sleep(5)
        
        fyil = int(browser.find_element_by_xpath("//*[@id='email-form']/div[3]/div[2]/div[1]/div[1]/div").text)
        time.sleep(2)
        
        if(fyil != yil):
            flager = fyil - yil
            if flager > 0:
                for i in range(flager):
                    cyil = browser.find_element_by_xpath('//*[@id="rightFinancialTableYearSliderButton"]/div')
                    cyil.click()
                    time.sleep(2)
            else:
                for i in range(abs(flager)):
                    cyil = browser.find_element_by_xpath('//*[@id="leftFinancialTableYearSliderButton"]/div')
                    cyil.click()
                    time.sleep(2)
        
        fdonem = 5 - donem  
        print(fdonem)           
        if(donem == 3 or donem == 4):
            while(fdonem > 0):
                cdonem = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]')
                cdonem.click()
                time.sleep(2)
                fdonem = fdonem - 1
        else:
            while(donem > 0):
                cdonem = browser.find_element_by_xpath('//*[@id="rightFinancialTablePeriodSliderButton"]')
                cdonem.click()
                time.sleep(2)
                donem = donem - 1
                
                
        getir = browser.find_element_by_xpath("//*[@id='Getir']")
        getir.click()
        time.sleep(5)

        try:  
            dosyaBulunamadi = browser.find_element_by_xpath("/html/body/div[10]/div/div/div[2]/div/div[2]")
            if dosyaBulunamadi:
                self.ui.bildirim.setText("Istedigin tarih ve doneme ait veriler bulunamadi!")
        except:
            self.ui.bildirim.setText("Istenilen tarih ve donemdeki tum sirketler cekildi!")       

    
    def sirketlerKap(self):
           
        options = webdriver.ChromeOptions() 
        adres = fileName2 +"\\Sirketler"
        #options.add_argument("download.default_directory="+ adres ")
        prefs = {
            "download.default_directory": adres,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
            }


        options.add_experimental_option('prefs', prefs)
        browser = webdriver.Chrome(chrome_options=options)

        browser.get("https://www.kap.org.tr/tr/api/exportCompanyPages/bist-sirketler/xls")
        time.sleep(20)

        browser.close()
        
        
        df_sirket = pd.read_html('Sirketler/Sirketler.xls')
        print(df_sirket)

        sirketler = []
        for i in range(len(df_sirket)):
             temp = df_sirket[i][1][1:]
             temp = temp.to_list()
             for k in range(len(temp)):
                 s = temp[k]
                 sirketler.append(s)
            
                
        model = QtGui.QStandardItemModel()
        self.ui.tumSirketler.setModel(model)
            
        for i in sirketler:
            item = QtGui.QStandardItem(i)
            model.appendRow(item)

    def widgetSelectedShow(self):
        self.ui.listWidget_2.clear()
#        items1 = self.ui.sirketler.selectedItems()
#        print(items1)
        items1 = [item.text() for item in self.ui.sirketler.selectedItems()]
        print(items1)    
        self.ui.listWidget_2.addItems(items1)
    
    def cekSecilen(self):
        lw = self.ui.listWidget_2
        items = []
        for x in range(lw.count()):
            items.append(str(lw.item(x).text()))
        print(items)


        a = 0         
        for sirketisim in items:
            passYap = False
            print(a)
            a = a + 1
            options = webdriver.ChromeOptions() 
            adres = fileName + "\Veriler\-"
            #options.add_argument("download.default_directory="+ adres ")
            prefs = {
            "download.default_directory": adres+sirketisim,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
            }
        
        
            options.add_experimental_option('prefs', prefs)
            browser = webdriver.Chrome(chrome_options=options)
        
            browser.get("https://www.kap.org.tr/tr/")
            time.sleep(5)
            

                        
            
            ftablolar = browser.find_element_by_xpath("//*[@id='financialTablesTab']/div")
            ftablolar.click()
        
            time.sleep(5)
            
            
            yilx = int(self.ui.lineEdit.text())
            
            fyil = int(browser.find_element_by_xpath("//*[@id='email-form']/div[3]/div[2]/div[1]/div[1]/div").text)
            print(fyil)
            print(sirketisim)
            time.sleep(2)
            if fyil == yilx:
                print(yilx)
            else:
                flager = fyil - yilx
                if flager > 0:
                    for i in range(flager):
                        cyil = browser.find_element_by_xpath('//*[@id="rightFinancialTableYearSliderButton"]/div')
                        cyil.click()
                        time.sleep(2)
                else:
                    for i in range(abs(flager)):
                        cyil = browser.find_element_by_xpath('//*[@id="leftFinancialTableYearSliderButton"]/div')
                        cyil.click()
                        time.sleep(2)
            
            
            try:
                sirket = browser.find_element_by_id("Sirket-6")
                sirket.send_keys(sirketisim)
                time.sleep(5)
                ftablolar2 = browser.find_element_by_xpath("//*[@id='calendarFilterInputFinancialTable']/div/a")
                ftablolar2.click()
                time.sleep(5)
            except:
                try:
                    sirket = browser.find_element_by_id("Sirket-6")
                    sirket.clear()
                    sirket.send_keys(sirketisim[:-1])
                    time.sleep(1)
                    ftablolar2 = browser.find_element_by_xpath("//*[@id='calendarFilterInputFinancialTable']/div/a")
                    ftablolar2.click()
                    time.sleep(1)
                except:
                    sirket = browser.find_element_by_id("Sirket-6")
                    sirket.clear()
                    sirket.send_keys(sirketisim)
                    time.sleep(1)
                    
              
              
                
            
            getir = browser.find_element_by_xpath("//*[@id='Getir']")
            getir.click()
            time.sleep(5)
            try:    
                dosyaBulunamadi = browser.find_element_by_xpath("/html/body/div[10]/div/div/div[2]/div/div[2]")
                
                if dosyaBulunamadi:
                    try:
                        solKaydir = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]/div') 
                        solKaydir.click()
                        solKaydir = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]/div') 
                        solKaydir.click()
                        time.sleep(2)
                        getir = browser.find_element_by_xpath("//*[@id='Getir']")
                        getir.click()
                        time.sleep(5)
                    except:
                        passYap == True
                        os.mkdir(adres+sirketisim)
                        print ("Successfully created the directory %s " % path)
            except:
                pass
            
            time.sleep(25)
        
        
            browser.close()
        
            if (path.exists(adres+sirketisim[:-1]+"\\2019-Tum Donemler.zip") == False) or (path.exists(adres+sirketisim+"\\2019-Tum Donemler.zip") == False):
                if passYap == True:    
                    self.ui.bildirim.setText("Tum veriler CEKILEMEDI!")
                    
                    break                    
                
        self.ui.bildirim.setText("Seçinler sirketler basariyla indirildi!")
    
    def seciliSec(self):
        
        print("ok")
    
      
    
    def bilancoExcel(self):
        
        sheets = pd.read_excel('Bilanco-Excel/Bilanco.xlsm' ,sheet_name=['KOZAA'])

        bilanco_isim = sheets['KOZAA'].iloc[:,0]
        bilanco_isim = bilanco_isim.values.tolist()
        #○bilanco_isim['bilanco'] = bilanco_isim['bilanco'].str.upper() 
        
        
        bilanco_isim_revize = []
        
        for i in bilanco_isim:
           
           if i[0] == ' ':
               new_i = list(i) 
               for letter in i:
                   if letter == ' ':
                       new_i.pop(0)
                   else:
                       i = (''.join(new_i))
                       bilanco_isim_revize.append(i.upper())
                       break
           else:
               bilanco_isim_revize.append(i.upper())
        
      
               
    
        print("Bitti !")
    
    def zipeAktar(self):
        self.ui.bosZipler.clear()
        
        veriler = os.listdir(fileName + "/Veriler/")
        bos_veri = []    
        for veri in veriler:
            path_sirket = []
            sirket = os.listdir(fileName2 +"\\Veriler\\"+veri)
            path_sirket.append(sirket)
            
        for zipex in veriler:
            path = fileName + "\\Veriler\\"
            path2 = zipex + "\\2019-Tum Donemler.zip"
            pathe = path + path2
            
        
            exact =fileName + "\\Excels"
           
            try:
                with ZipFile(pathe, 'r') as zipObj:
                    # Extract all the contents of zip file in current directory
                    zipObj.extractall(exact)
                print("ok")
            except:
                bos_veri.append(zipex)
                print("fail")
                  
        self.ui.bosZipler.addItems(bos_veri)
        
        
        
    def hepsiExcel(self):
        sheets = pd.read_excel('Bilanco-Excel/Bilanco.xlsm' ,sheet_name=['KOZAA'])

        bilanco_isim = sheets['KOZAA'].iloc[:,0]
        bilanco_isim = bilanco_isim.values.tolist()
        #○bilanco_isim['bilanco'] = bilanco_isim['bilanco'].str.upper() 
        excel_sheets = xlrd.open_workbook('Bilanco-Excel/Bilanco.xlsm', on_demand=True)
        excel_list = excel_sheets.sheet_names()
        excel_list.remove('Anatablo')
        excel_list.remove('HISSE-GRAFIK')
        excel_list.remove('GRAFİK 2')
        excel_list.remove('ÖZEL ORANLAR')
        excel_list.remove('Güncel Fiyat')
        excel_liste = [x.upper() for x in excel_list]
        print (excel_liste)
        
        
        cekSirkets = pd.read_excel('Hisseler/Hisseler.xlsx')
        cekSirketler = cekSirkets[["KOD"]].values.tolist()
        print(cekSirketler)
        
        bilanco_isim_revize = []
        
        for i in bilanco_isim:
           
           if i[0] == ' ':
               new_i = list(i) 
               for letter in i:
                   if letter == ' ':
                       new_i.pop(0)
                   else:
                       i = (''.join(new_i))
                       bilanco_isim_revize.append(i.upper())
                       break
           else:
               bilanco_isim_revize.append(i.upper())
               
        print(bilanco_isim_revize)
        
        excels = os.listdir("Excels/")

        matching = [s[:-4] for s in excels if '.xls' in s]
        
        print(len(matching))
        
        total = 0
        for excel in matching:
            temp = excel.split("-")
            keep = len(temp)
            total = total + keep
        print(total)
        
        npgive = np.empty([total,3], dtype = object)
        z = 0
        for i in range(len(matching)):
            temp = matching[i]
            x = temp.split("_")
            y = temp.split("-")
            for k in range(len(y)):
                if k == (len(y) - 1):
                    temp = y[-1].split("_")
                    npgive[z][0] = temp[0]
                    npgive[z][1] = x[-1]
                    npgive[z][2] = x[-2]
                    z += 1
                else:
                    npgive[z][0] = y[k]
                    npgive[z][1] = x[-1]
                    npgive[z][2] = x[-2]
                    z += 1
                    
        sirketKod =  pd.DataFrame({'Kod': npgive[:, 0], 'Donem': npgive[:, 1],'Yil': npgive[:, 2]})    
        print(sirketKod)
        
        yil = self.ui.lineEdit_3.text()
        donem = self.ui.lineEdit_4.text()
        print(yil)
        print(donem)
        yil = int(yil)
        donem = int(donem)
        donemlik = donem * 3
        
        is_sirketKod = sirketKod[(sirketKod.Yil == yil) & (sirketKod.Donem == donem)]
        
        print(is_sirketKod)
        
        olmadi = []
        a = 0
        b = 0
        for take in excel_liste:
            c = sirketKod[sirketKod.Kod == take.upper()]
            if(c.empty):
                print("fail")
                olmadi.append(take.upper())
            else:
                print("ok")
                
                b += 1
        print(olmadi)
        
        donemstr = str(donem)
        yilstr = str(yil)
        
        sonExcel = []
        for exc in matching:
            x = exc.split("_")
            if donemstr in x[-1] and yilstr in x[-2]:
                sonExcel.append(exc)
            else:
                continue
#        print(sonExcel)
        
        cekExcel = []
        for sExc in sonExcel:
            for excLi in cekSirketler:
                if excLi[0] in sExc:
                    cekExcel.append(sExc)
        
        cekexcel = [] 
        [cekexcel.append(x) for x in cekExcel if x not in cekexcel] 
        olmadis = []
        print(cekexcel)
        for excs in cekexcel:
            x = excs.split("-")
            if len(x) < 2:
                y = excs.split("_")
                print(excs)
                print(y[0])
                excs = str(excs) + ".xls"
                npsave = np.empty([len(bilanco_isim_revize),2], dtype = object)
        
        
                for i in range(len(bilanco_isim_revize)):
                    npsave[i][0] = bilanco_isim_revize[i]   
                    
                 
                
                #seçilen tablodan bilanço verilerinin ayıklanması       
                manu = pd.read_html("Excels/"+ excs)
                
                npsave[0][1] = str(yil) + "/" + str(donemlik)
                
                bilanchos = []
                for i in range(len(manu)):
                     if len(manu[i].columns) >= 5 and len(manu[i].columns) <= 8:
                         if len(manu[i])>2:
                             bilanchos.append(i)
                               
                newdf = manu[bilanchos[0]]        
                del bilanchos[0]
                
                newdf3 = manu[bilanchos[-1]]
                del bilanchos[-1]
                
                if len(manu[bilanchos[0]]) == 300:
                    newdf2 = manu[bilanchos[0]]
                    
                else:
                    frames = []
                    for i in range(len(bilanchos)):
                        frames.append(manu[bilanchos[i]])
                    if len(frames) == 0:
                        newdf2 = manu[bilanchos[0]]
                    elif len(frames) >= 1 :    
                        newdf2 = pd.concat(frames, ignore_index=True)    
                
                carpanx = manu[0]

                carpany = carpanx[1][0]
                
                carpanz  = carpany.strip(' TL')
                
                if not carpanz:
                    carpanz = 1
                else:
                    oldstr = carpanz
                    if isinstance(oldstr, int):
                        carpanz = oldstr
                    else:
                        newstr = oldstr.replace(".", "")
                        carpanz = int(newstr) 
                
                
                print(carpanz)
                
                
                print(len(newdf))
                print(len(newdf2))
                print(len(newdf3))
                
                for a in bilanchos:
                    print(len(manu[a]))
                
                
                #df1 için yapılması
                
                df1 = newdf[[1,3]].dropna(subset = [1])
                df1 = df1.reset_index()
                df1 = df1.drop("index",axis=1)
                df1 = df1.fillna(0)
                df1 = df1.reset_index()
                df1 = df1.drop("index",axis=1)
                df1 = df1.rename(columns={1: "bilanco", 3: "ciro"})
                df1['bilanco'] = df1['bilanco'].str.upper()
                df1 = df1.replace({'İ':'I'},regex = True)
                
                donen_varliklar =  df1.loc[2:54]
                ara_toplam_donenvarliklar = df1.loc[51].ciro
                toplam_donen_varlıklar = df1.loc[54].ciro 
                
                duran_varliklar = df1.loc[55:127]
                ozkaynak_yontemiyle_degerlenen_yatirimlar = df1.loc[68].ciro
                toplam_duran_varliklar = df1.loc[127].ciro
                
                toplam_varliklar = df1.loc[128].ciro
                
                kisa_vadeli_yukumlulukler = df1.loc[131:190]
                finansal_borclar = df1.loc[131].ciro
                diger_finansal_yukumlulukler = df1.loc[184].ciro
                musteri_soz_dogan_yuk = df1.loc[167].ciro
                ertelenmis_gelirler = df1.loc[176].ciro
                borc_karsiliklari = df1.loc[180].ciro
                ara_toplam_kisavadeliy = df1.loc[187].ciro
                toplam_kisa_vadeli = df1.loc[190].ciro
                
                uzun_vadeli_yukumlulukler = df1.loc[192:240]
                u_finansal_borclar = df1.loc[192].ciro
                u_musteri_soz_dogan_yuk = df1.loc[217].ciro
                u_ertelenmis_gelirler = df1.loc[226].ciro
                calisanlara_saglanan_faydalara = df1.loc[230].ciro
                toplam_uzun_vadeli = df1.loc[240].ciro
                
                
                ozkaynaklar = df1.loc[243:294]
                geçmis_yillar_kar_zararlari = df1.loc[291].ciro
                net_donem_kar_zaralari = df1.loc[292].ciro
                hisse_senedi_ihrac_primleri = df1.loc[251].ciro
                azinlik_paylari = df1.loc[293].ciro
                kalemler = df1.loc[245:281]
                kalemler = kalemler["ciro"].unique()
                diger_ozsermaye_kalemleri = 0
                for value in kalemler:
                    if value == 0:
                        topla = 0
                    else:    
                        topla = int(value.replace('.',''))
                    diger_ozsermaye_kalemleri = diger_ozsermaye_kalemleri + topla    
                toplam_ozkaynaklar = df1.loc[294].ciro
                
                toplam_kaynaklar = df1.loc[295].ciro
                
                
                
                
                
                
                for find in range(1,13):
                    cost = donen_varliklar[donen_varliklar["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                npsave[11][1] = int(ara_toplam_donenvarliklar.replace(".", ""))
                npsave[1][1] = int(toplam_donen_varlıklar.replace(".", ""))
                
                
                for find in range(13,30):
                    cost = duran_varliklar[duran_varliklar["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = ozkaynak_yontemiyle_degerlenen_yatirimlar
                if oldstr == 0:
                    npsave[19][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[19][1] = int(newstr)
                            
                npsave[13][1] = int(toplam_duran_varliklar.replace(".", ""))
                npsave[29][1] = int(toplam_varliklar.replace(".", ""))            
                            
                
                            
                for find in range(30,45):
                    cost = kisa_vadeli_yukumlulukler[kisa_vadeli_yukumlulukler["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = finansal_borclar
                if oldstr == 0:
                    npsave[32][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[32][1] = int(newstr)  
                
                oldstr = diger_finansal_yukumlulukler
                if oldstr == 0:
                    npsave[33][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[33][1] = int(newstr)  
                
                oldstr = musteri_soz_dogan_yuk
                if oldstr == 0:
                    npsave[36][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[36][1] = int(newstr)  
                
                oldstr = ertelenmis_gelirler
                if oldstr == 0:
                    npsave[39][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[39][1] = int(newstr) 
                
                oldstr = borc_karsiliklari
                if oldstr == 0:
                    npsave[41][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[41][1] = int(newstr) 
                
                
                npsave[43][1] = int(ara_toplam_kisavadeliy.replace(".", ""))
                npsave[31][1] = int(toplam_kisa_vadeli.replace(".", ""))    
                
                
                
                
                
                for find in range(45,58):
                    cost = uzun_vadeli_yukumlulukler[uzun_vadeli_yukumlulukler["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = u_finansal_borclar
                if oldstr == 0:
                    npsave[46][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[46][1] = int(newstr) 
                
                oldstr = u_musteri_soz_dogan_yuk
                if oldstr == 0:
                    npsave[50][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[50][1] = int(newstr) 
                    
                oldstr = u_ertelenmis_gelirler
                if oldstr == 0:
                    npsave[53][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[53][1] = int(newstr) 
                    
                oldstr = u_ertelenmis_gelirler
                if oldstr == 0:
                    npsave[53][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[53][1] = int(newstr) 
                    
                oldstr = calisanlara_saglanan_faydalara
                if oldstr == 0:
                    npsave[55][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[55][1] = int(newstr) 
                
                npsave[45][1] = int(toplam_uzun_vadeli.replace(".", ""))
                
                
                for find in range(58,71):
                    cost = ozkaynaklar[ozkaynaklar["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)            
                
                oldstr = geçmis_yillar_kar_zararlari
                if oldstr == 0:
                    npsave[66][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[66][1] = int(newstr) 
                
                oldstr = net_donem_kar_zaralari
                if oldstr == 0:
                    npsave[67][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[67][1] = int(newstr) 
                
                oldstr = hisse_senedi_ihrac_primleri
                if oldstr == 0:
                    npsave[62][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[62][1] = int(newstr) 
                
                oldstr = azinlik_paylari
                if oldstr == 0:
                    npsave[69][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[69][1] = int(newstr) 
                
                kalemler = df1.loc[245:281]
                kalemler = kalemler["ciro"].unique()
                diger_ozsermaye_kalemleri = 0
                for value in kalemler:
                    if value == 0:
                        topla = 0
                    else:    
                        topla = int(value.replace('.',''))
                    diger_ozsermaye_kalemleri = diger_ozsermaye_kalemleri + topla
                npsave[68][1] = diger_ozsermaye_kalemleri
                
                npsave[58][1] = int(toplam_ozkaynaklar.replace(".", ""))
                npsave[70][1] = int(toplam_kaynaklar.replace(".", "")) 
                    
                
                #df2 için yapılması
                
                
                df2 = newdf2[[1,3]].dropna(subset = [1])
                df2 = df2.reset_index()
                df2 = df2.drop("index",axis=1)
                df2 = df2.fillna(0)
                df2 = df2.reset_index()
                df2 = df2.drop("index",axis=1)
                df2 = df2.rename(columns={1: "bilanco", 3: "ciro"})
                df2['bilanco'] = df2['bilanco'].str.upper() 
                df2 = df2.replace({'İ':'I'},regex = True)
                
                surdurulen_faaliyetler= df2.loc[0:148]
                satis_gelirleri = df2.loc[2].ciro
                satislerin_maliyetleri = df2.loc[3].ciro
                f_u_p_k_diğer_ge = df2.loc[6].ciro
                f_u_p_k_diğer_gi = df2.loc[17].ciro    
                f_sektoru_faaliyetlerinden_diger_kar = df2.loc[15].ciro
                satis_diger_gelir_ve_giderler = df2.loc[27].ciro
                pazarlama_satis_ve_dagıtım_gider = df2.loc[32].ciro
                genel_yonetim_giderleri = df2.loc[31].ciro
                arastirma_ve_gelistirme_giderleri = df2.loc[33].ciro
                diger_faaliyet_gelirleri = df2.loc[34].ciro
                diger_faaliyet_giderleri = df2.loc[35].ciro
                faaliyet_kari_oncesi_diger_gelir_ve_giderl = df2.loc[36].ciro
                faaliyet_kari_zarari = df2.loc[37].ciro
                oldstr = faaliyet_kari_zarari
                if oldstr == 0:
                    a = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    a = int(newstr) 
                
                oldstr = diger_faaliyet_giderleri
                if oldstr == 0:
                    b = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    b = int(newstr) 
                    
                oldstr = diger_faaliyet_gelirleri
                if oldstr == 0:
                    c = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    c = int(newstr)     
             
                net_faaliyet_kar_zarari = a -( b + c)
                yatirim_faaliyetlerinden_giderler = df2.loc[41].ciro
                faaliyet_diger_gelir_ve_giderler = df2.loc[44].ciro
                ozkaynak_yontemiyle_degerlenen_yatırımlarin_kar_zarar = df2.loc[43].ciro
                finansman_gideri_oncesi_faaliyet_kari_zarari = df2.loc[48].ciro
                finansal_gelirler = df2.loc[49].ciro
                finansal_giderler = df2.loc[50].ciro
                surdurulen_faaliyetler_vergi_geliri = df2.loc[53].ciro
                donem_vergi_geliri = df2.loc[54].ciro
                ertelenmis_vergi_geliri = df2.loc[55].ciro
                surdurulen_faaliyetler_donem_kari_zarari = df2.loc[56].ciro
                durdurulan_faaliyetler_donem_kari_zarari = df2.loc[57].ciro
                durdurulan_faaliyetler_vergi_sonrasi_donem = df2.loc[57].ciro 
                azinlik_paylari = df2.loc[60].ciro
                
                
                for find in range(71,122):
                    cost = surdurulen_faaliyetler[surdurulen_faaliyetler["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = satis_gelirleri
                if oldstr == 0:
                    npsave[72][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[72][1] = int(newstr) 
                
                oldstr = satislerin_maliyetleri
                if oldstr == 0:
                    npsave[73][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[73][1] = int(newstr) 
                
                oldstr = f_u_p_k_diğer_ge
                if oldstr == 0:
                    npsave[76][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[76][1] = int(newstr) 
                
                oldstr = f_u_p_k_diğer_gi
                if oldstr == 0:
                    npsave[77][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[77][1] = int(newstr) 
                
                oldstr = f_sektoru_faaliyetlerinden_diger_kar
                if oldstr == 0:
                    npsave[78][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[78][1] = int(newstr) 
                
                oldstr = satis_diger_gelir_ve_giderler
                if oldstr == 0:
                    npsave[80][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[80][1] = int(newstr) 
                
                oldstr = pazarlama_satis_ve_dagıtım_gider
                if oldstr == 0:
                    npsave[82][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[82][1] = int(newstr) 
                
                oldstr = genel_yonetim_giderleri
                if oldstr == 0:
                    npsave[83][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[83][1] = int(newstr)
                
                oldstr = arastirma_ve_gelistirme_giderleri
                if oldstr == 0:
                    npsave[84][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[84][1] = int(newstr)
                
                oldstr = diger_faaliyet_gelirleri
                if oldstr == 0:
                    npsave[85][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[85][1] = int(newstr)
                
                oldstr = diger_faaliyet_giderleri
                if oldstr == 0:
                    npsave[86][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[86][1] = int(newstr)
                
                oldstr = faaliyet_kari_oncesi_diger_gelir_ve_giderl
                if oldstr == 0:
                    npsave[87][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[87][1] = int(newstr)
                
                oldstr = faaliyet_kari_zarari
                if oldstr == 0:
                    npsave[88][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[88][1] = int(newstr)
                
                
                oldstr = df2.loc[37].ciro
                if oldstr == 0:
                    a = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    a = int(newstr)
                
                oldstr = diger_faaliyet_giderleri
                if oldstr == 0:
                    b = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    b = int(newstr)
                    
                    
                oldstr = diger_faaliyet_gelirleri
                if oldstr == 0:
                    c = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    c = int(newstr) 
                    
                net_faaliyet_kar_zarari = a -( b + c)
                
                npsave[89][1] = net_faaliyet_kar_zarari
                
                oldstr = yatirim_faaliyetlerinden_giderler
                if oldstr == 0:
                    npsave[91][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[91][1] = int(newstr)
                
                oldstr = faaliyet_diger_gelir_ve_giderler
                if oldstr == 0:
                    npsave[92][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[92][1] = int(newstr)
                
                oldstr = ozkaynak_yontemiyle_degerlenen_yatırımlarin_kar_zarar
                if oldstr == 0:
                    npsave[93][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[93][1] = int(newstr)
                
                oldstr = finansman_gideri_oncesi_faaliyet_kari_zarari
                if oldstr == 0:
                    npsave[94][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[94][1] = int(newstr)
                
                oldstr = finansal_gelirler
                if oldstr == 0:
                    npsave[95][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[95][1] = int(newstr)
                
                oldstr = finansal_giderler
                if oldstr == 0:
                    npsave[96][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[96][1] = int(newstr)
                
                oldstr = surdurulen_faaliyetler_vergi_geliri
                if oldstr == 0:
                    npsave[99][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[99][1] = int(newstr)
                
                oldstr = donem_vergi_geliri
                if oldstr == 0:
                    npsave[100][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[100][1] = int(newstr)
                
                oldstr = ertelenmis_vergi_geliri
                if oldstr == 0:
                    npsave[101][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[101][1] = int(newstr)
                
                oldstr = surdurulen_faaliyetler_donem_kari_zarari
                if oldstr == 0:
                    npsave[103][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[103][1] = int(newstr)
                
                oldstr = durdurulan_faaliyetler_donem_kari_zarari
                if oldstr == 0:
                    npsave[106][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[106][1] = int(newstr)
                
                oldstr = durdurulan_faaliyetler_vergi_sonrasi_donem 
                if oldstr == 0:
                    npsave[105][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[105][1] = int(newstr)
                
                oldstr = azinlik_paylari
                if oldstr == 0:
                    npsave[108][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[108][1] = int(newstr)
                
                
                
                
                
                
                
                #df3 için yapılması
                
                df3 = newdf3[[1,3]].dropna(subset = [1])
                df3 = df3.reset_index()
                df3 = df3.drop("index",axis=1)
                df3 = df3.fillna(0)
                df3 = df3.reset_index()
                df3 = df3.drop("index",axis=1)
                df3 = df3.rename(columns={1: "bilanco", 3: "ciro"})
                df3['bilanco'] = df3['bilanco'].astype(str).str.upper() 
                df3 = df3.replace({'İ':'I'},regex = True)
                
                nakit_akislari = df3.loc[0:202]
                amortisman_giderleri = df3.loc[6].ciro
                
                
                npsave2 = np.empty([12,2],dtype = object)
                
                npsave2[0][0] = "IŞLETME FAALIYETLERINDEN NAKIT AKIŞLARI"
                npsave2[1][0] = "DÖNEM KARI (ZARARI)"
                npsave2[2][0] = "AMORTISMAN VE ITFA GIDERI ILE ILGILI DÜZELTMELER"
                npsave2[3][0] = "IŞLETME SERMAYESINDE GERÇEKLEŞEN DEĞIŞIMLER"
                npsave2[4][0] = "FINANSAL YATIRIMLARDAKI AZALIŞ (ARTIŞ)"
                npsave2[5][0] = "FAALIYETLERDEN ELDE EDILEN NAKIT AKIŞLARI"
                npsave2[6][0] = "YATIRIM FAALIYETLERINDEN KAYNAKLANAN NAKIT AKIŞLARI"
                npsave2[7][0] = "MADDI VE MADDI OLMAYAN DURAN VARLIKLARIN ALIMDAN KAYNAKLANAN NAKIT ÇIKIŞLARI"
                npsave2[8][0] = "FINANSMAN FAALIYETLERINDEN NAKIT AKIŞLARI"
                npsave2[9][0] = "NAKIT VE NAKIT BENZERLERINDEKI NET ARTIŞ (AZALIŞ)"
                npsave2[10][0] = "DÖNEM BAŞI NAKIT VE NAKIT BENZERLERI"
                npsave2[11][0] = "DÖNEM SONU NAKIT VE NAKIT BENZERLERI"
                
                
                for find in range(len(npsave2)):
                    cost = nakit_akislari[nakit_akislari["bilanco"] == npsave2[find][0]].ciro
                    if cost.empty:
                        npsave2[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave2[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave2[find][1] = int(newstr)
                
                
                
                sistem_1 = pd.DataFrame({'BİLANÇO': npsave[:, 0], 'CİRO': npsave[:, 1]})
                sistem_2 = pd.DataFrame({'BİLANÇO': npsave2[:, 0], 'CİRO': npsave2[:, 1]})
                
                excel_aktar = sistem_1.append(sistem_2, ignore_index = True)
                
                excel_aktar["CIRO"] = excel_aktar["CİRO"] * carpanz
                
                app = xw.App(visible=False) # IF YOU WANT EXCEL TO RUN IN BACKGROUND
                
                xlwb = xw.Book('Bilanco-Excel/Bilanco.xlsm')
                
                try:
                    xlws = xlwb.sheets[y[0].upper()]
                except:
                    try:
                        xlws = xlwb.sheets[y[0].lower()]
                    except:
                        xlwb.close()
                        app.kill()
                        olmadis.append(y[0])
                        continue
                    
                xlws.range("B:B").insert('right')
                donem = list(excel_aktar.CİRO)
                xlws.range('B2').value = donem[0]
                ciro = list(excel_aktar.CIRO)
                xlws.range('B3').options(transpose=True).value = ciro[1:]
                
                xlwb.save()
                xlwb.close()
                app.kill()

            
            else:
                y = excs.split("_")
                z = y[0].split("-")
                print(excs)
                
                excs = str(excs) + ".xls"
                
            
                npsave = np.empty([len(bilanco_isim_revize),2], dtype = object)
        
        
                for i in range(len(bilanco_isim_revize)):
                    npsave[i][0] = bilanco_isim_revize[i]   
                    
                 
                
                #seçilen tablodan bilanço verilerinin ayıklanması       
                manu = pd.read_html("Excels/"+ excs)
                
                npsave[0][1] = str(yil) + "/" + str(donemlik)
                
                bilanchos = []
                for i in range(len(manu)):
                     if len(manu[i].columns) >= 5 and len(manu[i].columns) <= 8:
                         if len(manu[i])>2:
                             bilanchos.append(i)
                               
                newdf = manu[bilanchos[0]]
                del bilanchos[0]
                
                newdf3 = manu[bilanchos[-1]]
                del bilanchos[-1]
                
                if len(manu[bilanchos[0]]) == 300:
                    newdf2 = manu[bilanchos[0]]
                    
                else:
                    frames = []
                    for i in range(len(bilanchos)):
                        frames.append(manu[bilanchos[i]])
                        
                    if len(frames) == 1:
                        newdf2 = manu[bilanchos[0]]
                    elif len(frames) >= 1 :    
                        newdf2 = pd.concat(frames, ignore_index=True)   
                
                
                carpanx = manu[0]

                carpany = carpanx[1][0]
                
                carpanz  = carpany.strip(' TL')
                
                if not carpanz:
                    carpanz = 1
                else:
                    oldstr = carpanz
                    if isinstance(oldstr, int):
                        carpanz = oldstr
                    else:
                        newstr = oldstr.replace(".", "")
                        carpanz = int(newstr) 
                
                print(carpanz)
              
                
                print(len(newdf))
                print(len(newdf2))
                print(len(newdf3))
                for a in bilanchos:
                    print(len(manu[a]))
                #df1 için yapılması
                
                df1 = newdf[[1,3]].dropna(subset = [1])
                df1 = df1.reset_index()
                df1 = df1.drop("index",axis=1)
                df1 = df1.fillna(0)
                df1 = df1.reset_index()
                df1 = df1.drop("index",axis=1)
                df1 = df1.rename(columns={1: "bilanco", 3: "ciro"})
                df1['bilanco'] = df1['bilanco'].str.upper()
                df1 = df1.replace({'İ':'I'},regex = True)
                
                donen_varliklar =  df1.loc[2:54]
                ara_toplam_donenvarliklar = df1.loc[51].ciro
                toplam_donen_varlıklar = df1.loc[54].ciro 
                
                duran_varliklar = df1.loc[55:127]
                ozkaynak_yontemiyle_degerlenen_yatirimlar = df1.loc[68].ciro
                toplam_duran_varliklar = df1.loc[127].ciro
                
                toplam_varliklar = df1.loc[128].ciro
                
                kisa_vadeli_yukumlulukler = df1.loc[131:190]
                finansal_borclar = df1.loc[131].ciro
                diger_finansal_yukumlulukler = df1.loc[184].ciro
                musteri_soz_dogan_yuk = df1.loc[167].ciro
                ertelenmis_gelirler = df1.loc[176].ciro
                borc_karsiliklari = df1.loc[180].ciro
                ara_toplam_kisavadeliy = df1.loc[187].ciro
                toplam_kisa_vadeli = df1.loc[190].ciro
                
                uzun_vadeli_yukumlulukler = df1.loc[192:240]
                u_finansal_borclar = df1.loc[192].ciro
                u_musteri_soz_dogan_yuk = df1.loc[217].ciro
                u_ertelenmis_gelirler = df1.loc[226].ciro
                calisanlara_saglanan_faydalara = df1.loc[230].ciro
                toplam_uzun_vadeli = df1.loc[240].ciro
                
                
                ozkaynaklar = df1.loc[243:294]
                geçmis_yillar_kar_zararlari = df1.loc[291].ciro
                net_donem_kar_zaralari = df1.loc[292].ciro
                hisse_senedi_ihrac_primleri = df1.loc[251].ciro
                azinlik_paylari = df1.loc[293].ciro
                kalemler = df1.loc[245:281]
                kalemler = kalemler["ciro"].unique()
                diger_ozsermaye_kalemleri = 0
                for value in kalemler:
                    if value == 0:
                        topla = 0
                    else:    
                        topla = int(value.replace('.',''))
                    diger_ozsermaye_kalemleri = diger_ozsermaye_kalemleri + topla    
                toplam_ozkaynaklar = df1.loc[294].ciro
                
                toplam_kaynaklar = df1.loc[295].ciro
                
                
                
                
                
                
                for find in range(1,13):
                    cost = donen_varliklar[donen_varliklar["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                npsave[11][1] = int(ara_toplam_donenvarliklar.replace(".", ""))
                npsave[1][1] = int(toplam_donen_varlıklar.replace(".", ""))
                
                
                for find in range(13,30):
                    cost = duran_varliklar[duran_varliklar["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = ozkaynak_yontemiyle_degerlenen_yatirimlar
                if oldstr == 0:
                    npsave[19][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[19][1] = int(newstr)
                            
                npsave[13][1] = int(toplam_duran_varliklar.replace(".", ""))
                npsave[29][1] = int(toplam_varliklar.replace(".", ""))            
                            
                
                            
                for find in range(30,45):
                    cost = kisa_vadeli_yukumlulukler[kisa_vadeli_yukumlulukler["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = finansal_borclar
                if oldstr == 0:
                    npsave[32][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[32][1] = int(newstr)  
                
                oldstr = diger_finansal_yukumlulukler
                if oldstr == 0:
                    npsave[33][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[33][1] = int(newstr)  
                
                oldstr = musteri_soz_dogan_yuk
                if oldstr == 0:
                    npsave[36][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[36][1] = int(newstr)  
                
                oldstr = ertelenmis_gelirler
                if oldstr == 0:
                    npsave[39][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[39][1] = int(newstr) 
                
                oldstr = borc_karsiliklari
                if oldstr == 0:
                    npsave[41][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[41][1] = int(newstr) 
                
                
                npsave[43][1] = int(ara_toplam_kisavadeliy.replace(".", ""))
                npsave[31][1] = int(toplam_kisa_vadeli.replace(".", ""))    
                
                
                
                
                
                for find in range(45,58):
                    cost = uzun_vadeli_yukumlulukler[uzun_vadeli_yukumlulukler["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = u_finansal_borclar
                if oldstr == 0:
                    npsave[46][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[46][1] = int(newstr) 
                
                oldstr = u_musteri_soz_dogan_yuk
                if oldstr == 0:
                    npsave[50][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[50][1] = int(newstr) 
                    
                oldstr = u_ertelenmis_gelirler
                if oldstr == 0:
                    npsave[53][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[53][1] = int(newstr) 
                    
                oldstr = u_ertelenmis_gelirler
                if oldstr == 0:
                    npsave[53][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[53][1] = int(newstr) 
                    
                oldstr = calisanlara_saglanan_faydalara
                if oldstr == 0:
                    npsave[55][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[55][1] = int(newstr) 
                
                npsave[45][1] = int(toplam_uzun_vadeli.replace(".", ""))
                
                
                for find in range(58,71):
                    cost = ozkaynaklar[ozkaynaklar["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)            
                
                oldstr = geçmis_yillar_kar_zararlari
                if oldstr == 0:
                    npsave[66][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[66][1] = int(newstr) 
                
                oldstr = net_donem_kar_zaralari
                if oldstr == 0:
                    npsave[67][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[67][1] = int(newstr) 
                
                oldstr = hisse_senedi_ihrac_primleri
                if oldstr == 0:
                    npsave[62][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[62][1] = int(newstr) 
                
                oldstr = azinlik_paylari
                if oldstr == 0:
                    npsave[69][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[69][1] = int(newstr) 
                
                kalemler = df1.loc[245:281]
                kalemler = kalemler["ciro"].unique()
                diger_ozsermaye_kalemleri = 0
                for value in kalemler:
                    if value == 0:
                        topla = 0
                    else:    
                        topla = int(value.replace('.',''))
                    diger_ozsermaye_kalemleri = diger_ozsermaye_kalemleri + topla
                npsave[68][1] = diger_ozsermaye_kalemleri
                
                npsave[58][1] = int(toplam_ozkaynaklar.replace(".", ""))
                npsave[70][1] = int(toplam_kaynaklar.replace(".", "")) 
                    
                
                #df2 için yapılması
                
                
                df2 = newdf2[[1,3]].dropna(subset = [1])
                df2 = df2.reset_index()
                df2 = df2.drop("index",axis=1)
                df2 = df2.fillna(0)
                df2 = df2.reset_index()
                df2 = df2.drop("index",axis=1)
                df2 = df2.rename(columns={1: "bilanco", 3: "ciro"})
                df2['bilanco'] = df2['bilanco'].str.upper() 
                df2 = df2.replace({'İ':'I'},regex = True)
                
                surdurulen_faaliyetler= df2.loc[0:148]
                satis_gelirleri = df2.loc[2].ciro
                satislerin_maliyetleri = df2.loc[3].ciro
                f_u_p_k_diğer_ge = df2.loc[6].ciro
                f_u_p_k_diğer_gi = df2.loc[17].ciro    
                f_sektoru_faaliyetlerinden_diger_kar = df2.loc[15].ciro
                satis_diger_gelir_ve_giderler = df2.loc[27].ciro
                pazarlama_satis_ve_dagıtım_gider = df2.loc[32].ciro
                genel_yonetim_giderleri = df2.loc[31].ciro
                arastirma_ve_gelistirme_giderleri = df2.loc[33].ciro
                diger_faaliyet_gelirleri = df2.loc[34].ciro
                diger_faaliyet_giderleri = df2.loc[35].ciro
                faaliyet_kari_oncesi_diger_gelir_ve_giderl = df2.loc[36].ciro
                faaliyet_kari_zarari = df2.loc[37].ciro
                oldstr = faaliyet_kari_zarari
                if oldstr == 0:
                    a = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    a = int(newstr) 
                
                oldstr = diger_faaliyet_giderleri
                if oldstr == 0:
                    b = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    b = int(newstr) 
                    
                oldstr = diger_faaliyet_gelirleri
                if oldstr == 0:
                    c = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    c = int(newstr)     
                net_faaliyet_kar_zarari = a -( b + c)
                yatirim_faaliyetlerinden_giderler = df2.loc[41].ciro
                faaliyet_diger_gelir_ve_giderler = df2.loc[44].ciro
                ozkaynak_yontemiyle_degerlenen_yatırımlarin_kar_zarar = df2.loc[43].ciro
                finansman_gideri_oncesi_faaliyet_kari_zarari = df2.loc[48].ciro
                finansal_gelirler = df2.loc[49].ciro
                finansal_giderler = df2.loc[50].ciro
                surdurulen_faaliyetler_vergi_geliri = df2.loc[53].ciro
                donem_vergi_geliri = df2.loc[54].ciro
                ertelenmis_vergi_geliri = df2.loc[55].ciro
                surdurulen_faaliyetler_donem_kari_zarari = df2.loc[56].ciro
                durdurulan_faaliyetler_donem_kari_zarari = df2.loc[57].ciro
                durdurulan_faaliyetler_vergi_sonrasi_donem = df2.loc[57].ciro 
                azinlik_paylari = df2.loc[60].ciro
                
                
                for find in range(71,122):
                    cost = surdurulen_faaliyetler[surdurulen_faaliyetler["bilanco"] == npsave[find][0]].ciro
                    if cost.empty:
                        npsave[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave[find][1] = int(newstr)
                
                oldstr = satis_gelirleri
                if oldstr == 0:
                    npsave[72][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[72][1] = int(newstr) 
                
                oldstr = satislerin_maliyetleri
                if oldstr == 0:
                    npsave[73][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[73][1] = int(newstr) 
                
                oldstr = f_u_p_k_diğer_ge
                if oldstr == 0:
                    npsave[76][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[76][1] = int(newstr) 
                
                oldstr = f_u_p_k_diğer_gi
                if oldstr == 0:
                    npsave[77][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[77][1] = int(newstr) 
                
                oldstr = f_sektoru_faaliyetlerinden_diger_kar
                if oldstr == 0:
                    npsave[78][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[78][1] = int(newstr) 
                
                oldstr = satis_diger_gelir_ve_giderler
                if oldstr == 0:
                    npsave[80][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[80][1] = int(newstr) 
                
                oldstr = pazarlama_satis_ve_dagıtım_gider
                if oldstr == 0:
                    npsave[82][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[82][1] = int(newstr) 
                
                oldstr = genel_yonetim_giderleri
                if oldstr == 0:
                    npsave[83][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[83][1] = int(newstr)
                
                oldstr = arastirma_ve_gelistirme_giderleri
                if oldstr == 0:
                    npsave[84][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[84][1] = int(newstr)
                
                oldstr = diger_faaliyet_gelirleri
                if oldstr == 0:
                    npsave[85][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[85][1] = int(newstr)
                
                oldstr = diger_faaliyet_giderleri
                if oldstr == 0:
                    npsave[86][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[86][1] = int(newstr)
                
                oldstr = faaliyet_kari_oncesi_diger_gelir_ve_giderl
                if oldstr == 0:
                    npsave[87][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[87][1] = int(newstr)
                
                oldstr = faaliyet_kari_zarari
                if oldstr == 0:
                    npsave[88][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[88][1] = int(newstr)
                
                
                oldstr = df2.loc[37].ciro
                if oldstr == 0:
                    a = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    a = int(newstr)
                
                oldstr = diger_faaliyet_giderleri
                if oldstr == 0:
                    b = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    b = int(newstr)
                    
                    
                oldstr = diger_faaliyet_gelirleri
                if oldstr == 0:
                    c = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    c = int(newstr)    
                
                
                net_faaliyet_kar_zarari = a -( b + c)
                
                npsave[89][1] = net_faaliyet_kar_zarari
                
                oldstr = yatirim_faaliyetlerinden_giderler
                if oldstr == 0:
                    npsave[91][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[91][1] = int(newstr)
                
                oldstr = faaliyet_diger_gelir_ve_giderler
                if oldstr == 0:
                    npsave[92][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[92][1] = int(newstr)
                
                oldstr = ozkaynak_yontemiyle_degerlenen_yatırımlarin_kar_zarar
                if oldstr == 0:
                    npsave[93][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[93][1] = int(newstr)
                
                oldstr = finansman_gideri_oncesi_faaliyet_kari_zarari
                if oldstr == 0:
                    npsave[94][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[94][1] = int(newstr)
                
                oldstr = finansal_gelirler
                if oldstr == 0:
                    npsave[95][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[95][1] = int(newstr)
                
                oldstr = finansal_giderler
                if oldstr == 0:
                    npsave[96][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[96][1] = int(newstr)
                
                oldstr = surdurulen_faaliyetler_vergi_geliri
                if oldstr == 0:
                    npsave[99][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[99][1] = int(newstr)
                
                oldstr = donem_vergi_geliri
                if oldstr == 0:
                    npsave[100][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[100][1] = int(newstr)
                
                oldstr = ertelenmis_vergi_geliri
                if oldstr == 0:
                    npsave[101][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[101][1] = int(newstr)
                
                oldstr = surdurulen_faaliyetler_donem_kari_zarari
                if oldstr == 0:
                    npsave[103][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[103][1] = int(newstr)
                
                oldstr = durdurulan_faaliyetler_donem_kari_zarari
                if oldstr == 0:
                    npsave[106][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[106][1] = int(newstr)
                
                oldstr = durdurulan_faaliyetler_vergi_sonrasi_donem 
                if oldstr == 0:
                    npsave[105][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[105][1] = int(newstr)
                
                oldstr = azinlik_paylari
                if oldstr == 0:
                    npsave[108][1] = oldstr
                else:
                    newstr = oldstr.replace(".", "")
                    npsave[108][1] = int(newstr)
                
                
                
                
                
                
                
                #df3 için yapılması
                
                df3 = newdf3[[1,3]].dropna(subset = [1])
                df3 = df3.reset_index()
                df3 = df3.drop("index",axis=1)
                df3 = df3.fillna(0)
                df3 = df3.reset_index()
                df3 = df3.drop("index",axis=1)
                df3 = df3.rename(columns={1: "bilanco", 3: "ciro"})
                df3['bilanco'] = df3['bilanco'].astype(str).str.upper() 
                df3 = df3.replace({'İ':'I'},regex = True)
                
                nakit_akislari = df3.loc[0:202]
                amortisman_giderleri = df3.loc[6].ciro
                
                
                npsave2 = np.empty([12,2],dtype = object)
                
                npsave2[0][0] = "IŞLETME FAALIYETLERINDEN NAKIT AKIŞLARI"
                npsave2[1][0] = "DÖNEM KARI (ZARARI)"
                npsave2[2][0] = "AMORTISMAN VE ITFA GIDERI ILE ILGILI DÜZELTMELER"
                npsave2[3][0] = "IŞLETME SERMAYESINDE GERÇEKLEŞEN DEĞIŞIMLER"
                npsave2[4][0] = "FINANSAL YATIRIMLARDAKI AZALIŞ (ARTIŞ)"
                npsave2[5][0] = "FAALIYETLERDEN ELDE EDILEN NAKIT AKIŞLARI"
                npsave2[6][0] = "YATIRIM FAALIYETLERINDEN KAYNAKLANAN NAKIT AKIŞLARI"
                npsave2[7][0] = "MADDI VE MADDI OLMAYAN DURAN VARLIKLARIN ALIMDAN KAYNAKLANAN NAKIT ÇIKIŞLARI"
                npsave2[8][0] = "FINANSMAN FAALIYETLERINDEN NAKIT AKIŞLARI"
                npsave2[9][0] = "NAKIT VE NAKIT BENZERLERINDEKI NET ARTIŞ (AZALIŞ)"
                npsave2[10][0] = "DÖNEM BAŞI NAKIT VE NAKIT BENZERLERI"
                npsave2[11][0] = "DÖNEM SONU NAKIT VE NAKIT BENZERLERI"
                
                
                for find in range(len(npsave2)):
                    cost = nakit_akislari[nakit_akislari["bilanco"] == npsave2[find][0]].ciro
                    if cost.empty:
                        npsave2[find][1] = 0
                    else:
                        oldstr = cost.iloc[0]
                        if oldstr == 0:
                            npsave2[find][1] = oldstr
                        else:
                            newstr = oldstr.replace(".", "")
                            npsave2[find][1] = int(newstr)
                
                
                
                sistem_1 = pd.DataFrame({'BİLANÇO': npsave[:, 0], 'CİRO': npsave[:, 1]})
                sistem_2 = pd.DataFrame({'BİLANÇO': npsave2[:, 0], 'CİRO': npsave2[:, 1]})
                
                excel_aktar = sistem_1.append(sistem_2, ignore_index = True)
                
                excel_aktar["CIRO"] = excel_aktar["CİRO"] * carpanz
                   
                for items in z:
                    print(items)
                    
                    app = xw.App(visible=False) # IF YOU WANT EXCEL TO RUN IN BACKGROUND
                    
                    xlwb = xw.Book('Bilanco-Excel/Bilanco.xlsm')
                    
                    try:
                        xlws = xlwb.sheets[items.upper()]
                    except:
                        try:
                            xlws = xlwb.sheets[items.lower()]
                        except:
                            xlwb.close()
                            app.kill()
                            olmadis.append(items)
                            continue
    
                    
                        
                        
                    xlws.range("B:B").insert('right')
                    donem = list(excel_aktar.CİRO)
                    xlws.range('B2').value = donem[0]
                    ciro = list(excel_aktar.CIRO)
                    xlws.range('B3').options(transpose=True).value = ciro[1:]
                    
                    xlwb.save()
                    xlwb.close()
                    app.kill()
        
        self.ui.listWidget.addItems(olmadis)        
        self.ui.bildirim.setText("Veriler excel'e aktarildi!")        
            
    def listeyeDok(self):
        
        
        df_sirket = pd.read_html('Sirketler/Sirketler.xls')
        print(df_sirket)

        sirketler = []
        for i in range(len(df_sirket)):
             temp = df_sirket[i][1][1:]
             temp = temp.to_list()
             for k in range(len(temp)):
                 s = temp[k]
                 sirketler.append(s)
            
                
        model = QtGui.QStandardItemModel()
        self.ui.tumSirketler.setModel(model)
            
        for i in sirketler:
            item = QtGui.QStandardItem(i)
            model.appendRow(item)
        
        
        
        self.ui.bildirim.setText("Sirket Verileri Cekildi!")        
            # self.gridLayout.addWidget(self.listView, 1, 0, 1, 2)
    
    def widgetListele(self):
        
        self.ui.sirketler.clear()
        
        df_sirket = pd.read_html('Sirketler/Sirketler.xls')
        print(df_sirket)

        sirketler = []
        for i in range(len(df_sirket)):
             temp = df_sirket[i][1][1:]
             temp = temp.to_list()
             for k in range(len(temp)):
                 s = temp[k]
                 sirketler.append(s)
                 
       
        veriler = os.listdir(fileName + "/Veriler/")
        
        alinmis_Sirketler = []
        for alinmis in veriler:
            gom = alinmis[1:] + "."
            alinmis_Sirketler.append(gom)
        
        sonListe = []
        for sirket in sirketler:           
            if str(sirket + ".") in alinmis_Sirketler:
                continue
            elif sirket not in alinmis_Sirketler:
                sonListe.append(sirket)
                
                 
        self.ui.sirketler.addItems(sonListe)

                  
        
        
        
       
        
        
    
    def genelYukle(self):
        df_sirket = pd.read_html('Sirketler/Sirketler.xls')
        print(df_sirket)

        sirketler = []
        for i in range(len(df_sirket)):
             temp = df_sirket[i][1][1:]
             temp = temp.to_list()
             for k in range(len(temp)):
                 s = temp[k]
                 sirketler.append(s)
        
        a = 0         
        for sirketisim in sirketler:
            passYap = False
            print(a)
            a = a + 1
            options = webdriver.ChromeOptions() 
            adres = fileName + "\Veriler\-"
            #options.add_argument("download.default_directory="+ adres ")
            prefs = {
            "download.default_directory": adres+sirketisim,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
            }
        
        
            options.add_experimental_option('prefs', prefs)
            browser = webdriver.Chrome(chrome_options=options)
        
            browser.get("https://www.kap.org.tr/tr/")
            time.sleep(5)
            
            
            
            ftablolar = browser.find_element_by_xpath("//*[@id='financialTablesTab']/div")
            ftablolar.click()
        
            time.sleep(5)
            
            
                        
            yilx = int(self.ui.lineEdit.text())
            
            fyil = int(browser.find_element_by_xpath("//*[@id='email-form']/div[3]/div[2]/div[1]/div[1]/div").text)
            print(fyil)
            print(sirketisim)
            time.sleep(2)
            if fyil == yilx:
                print(yilx)
            else:
                flager = fyil - yilx
                if flager > 0:
                    for i in range(flager):
                        cyil = browser.find_element_by_xpath('//*[@id="rightFinancialTableYearSliderButton"]/div')
                        cyil.click()
                        time.sleep(2)
                else:
                    for i in range(abs(flager)):
                        cyil = browser.find_element_by_xpath('//*[@id="leftFinancialTableYearSliderButton"]/div')
                        cyil.click()
                        time.sleep(2)
            
            try:
                sirket = browser.find_element_by_id("Sirket-6")
                sirket.send_keys(sirketisim)
                time.sleep(5)
                ftablolar2 = browser.find_element_by_xpath("//*[@id='calendarFilterInputFinancialTable']/div/a")
                ftablolar2.click()
                time.sleep(5)
            except:
                try:
                    sirket = browser.find_element_by_id("Sirket-6")
                    sirket.clear()
                    sirket.send_keys(sirketisim[:-1])
                    time.sleep(1)
                    ftablolar2 = browser.find_element_by_xpath("//*[@id='calendarFilterInputFinancialTable']/div/a")
                    ftablolar2.click()
                    time.sleep(1)
                except:
                    sirket = browser.find_element_by_id("Sirket-6")
                    sirket.clear()
                    sirket.send_keys(sirketisim)
                    time.sleep(1)
                    
              
              
                
            
            getir = browser.find_element_by_xpath("//*[@id='Getir']")
            getir.click()
            time.sleep(5)
            try:    
                dosyaBulunamadi = browser.find_element_by_xpath("/html/body/div[10]/div/div/div[2]/div/div[2]")
                
                if dosyaBulunamadi:
                    try:
                        solKaydir = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]/div') 
                        solKaydir.click()
                        solKaydir = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]/div') 
                        solKaydir.click()
                        time.sleep(2)
                        getir = browser.find_element_by_xpath("//*[@id='Getir']")
                        getir.click()
                        time.sleep(5)
                    except:
                        passYap == True
                        os.mkdir(adres+sirketisim)
                        print ("Successfully created the directory %s " % path)
            except:
                pass
            
            time.sleep(25)
        
        
            browser.close()
        
            if (path.exists(adres+sirketisim[:-1]+"\\2019-Tum Donemler.zip") == False) or (path.exists(adres+sirketisim+"\\2019-Tum Donemler.zip") == False):
                if passYap == True:    
                    self.ui.bildirim.setText("Tum veriler CEKILEMEDI!")
                    
                    break                    
                
        self.ui.bildirim.setText("Seçinler sirketler basariyla indirildi!")
 
            
            
            

    def silYedekle(self):

        today = date.today()
        shutil.copy('Bilanco-Excel/Bilanco.xlsm', 'BilancoYedek/BilancoBackUp-'+str(today)+'.xlsm')
        
        
        #Excels ve Veriler Dosyalarının İçindki Tüm Verileri Siler
        silveri = os.listdir("Veriler/")
        silveri2 = os.listdir("Excels/")
        
        try:
            for veri in silveri:
                path = "Veriler/"+veri
                shutil.rmtree(path)
            
            for veri2 in silveri2:
                try:
                    path = "Excels/"+veri2
                    os.remove(path)
                except:
                    path = "Excels/"+veri2
                    shutil.rmtree(path)
        except:
            pass
        
        
        
        silveri3 = os.listdir("Sirketler/")
        
        try:
                
            for veri3 in silveri3:
                try:
                    path = "Sirketler/"+veri3
                    os.remove(path)
                except:
                    path = "Sirketler/"+veri3
                    shutil.rmtree(path)
        except:
            pass
        
    def devamEttir(self):
        
        df_sirket = pd.read_html('Sirketler/Sirketler.xls')
        print(df_sirket)

        sirketler = []
        for i in range(len(df_sirket)):
             temp = df_sirket[i][1][1:]
             temp = temp.to_list()
             for k in range(len(temp)):
                 s = temp[k]
                 sirketler.append(s)
                 
       
        veriler = os.listdir(fileName + "/Veriler/")
        
        alinmis_Sirketler = []
        for alinmis in veriler:
            gom = alinmis[1:] + "."
            alinmis_Sirketler.append(gom)
        
        sonListe = []
        for sirket in sirketler:
            if str(sirket + ".") in alinmis_Sirketler:
                continue
            elif sirket not in alinmis_Sirketler:
                sonListe.append(sirket)   
                
        a = 0         
        for sirketisim in sonListe:
            passYap = False
            print(a)
            a = a + 1
            options = webdriver.ChromeOptions() 
            adres = fileName + "\Veriler\-"
            #options.add_argument("download.default_directory="+ adres ")
            prefs = {
            "download.default_directory": adres+sirketisim,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
            }
        
        
            options.add_experimental_option('prefs', prefs)
            browser = webdriver.Chrome(chrome_options=options)
        
            browser.get("https://www.kap.org.tr/tr/")
            time.sleep(5)
            
            
            
            ftablolar = browser.find_element_by_xpath("//*[@id='financialTablesTab']/div")
            ftablolar.click()
        
            time.sleep(5)
            
            
                        
            yilx = int(self.ui.lineEdit.text())
            
            fyil = int(browser.find_element_by_xpath("//*[@id='email-form']/div[3]/div[2]/div[1]/div[1]/div").text)
            print(fyil)
            print(sirketisim)
            time.sleep(2)
            if fyil == yilx:
                print(yilx)
            else:
                flager = fyil - yilx
                if flager > 0:
                    for i in range(flager):
                        cyil = browser.find_element_by_xpath('//*[@id="rightFinancialTableYearSliderButton"]/div')
                        cyil.click()
                        time.sleep(2)
                else:
                    for i in range(abs(flager)):
                        cyil = browser.find_element_by_xpath('//*[@id="leftFinancialTableYearSliderButton"]/div')
                        cyil.click()
                        time.sleep(2)
            
            
            try:
                sirket = browser.find_element_by_id("Sirket-6")
                sirket.send_keys(sirketisim)
                time.sleep(5)
                ftablolar2 = browser.find_element_by_xpath("//*[@id='calendarFilterInputFinancialTable']/div/a")
                ftablolar2.click()
                time.sleep(5)
            except:
                try:
                    sirket = browser.find_element_by_id("Sirket-6")
                    sirket.clear()
                    sirket.send_keys(sirketisim[:-1])
                    time.sleep(1)
                    ftablolar2 = browser.find_element_by_xpath("//*[@id='calendarFilterInputFinancialTable']/div/a")
                    ftablolar2.click()
                    time.sleep(1)
                except:
                    sirket = browser.find_element_by_id("Sirket-6")
                    sirket.clear()
                    sirket.send_keys(sirketisim)
                    time.sleep(1)
                    
              
              
                
            
            getir = browser.find_element_by_xpath("//*[@id='Getir']")
            getir.click()
            time.sleep(5)
            try:    
                dosyaBulunamadi = browser.find_element_by_xpath("/html/body/div[10]/div/div/div[2]/div/div[2]")
                
                if dosyaBulunamadi:
                    try:
                        solKaydir = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]/div') 
                        solKaydir.click()
                        solKaydir = browser.find_element_by_xpath('//*[@id="leftFinancialTablePeriodSliderButton"]/div') 
                        solKaydir.click()
                        time.sleep(2)
                        getir = browser.find_element_by_xpath("//*[@id='Getir']")
                        getir.click()
                        time.sleep(5)
                    except:
                        passYap == True
                        os.mkdir(adres+sirketisim)
                        print ("Successfully created the directory %s " % path)
            except:
                pass
            
            time.sleep(25)
        
        
            browser.close()
        
            if (path.exists(adres+sirketisim[:-1]+"\\2019-Tum Donemler.zip") == False) or (path.exists(adres+sirketisim+"\\2019-Tum Donemler.zip") == False):
                if passYap == True:    
                    self.ui.bildirim.setText("Tum veriler CEKILEMEDI!")
                    
                    break                    
                
        self.ui.bildirim.setText("Seçinler sirketler basariyla indirildi!")



if __name__ == "__main__":
    import sys
    import os
    import html5lib
    fileName = os.path.abspath(os.getcwd())
    fileName2 = fileName
    fileName2= str(fileName2.replace('/', '\\'))
    print(fileName)
    print(fileName2)
    app = QApplication(sys.argv)
    MainWindow = Bilanco()
    MainWindow.show()
    sys.exit(app.exec_())
    
    
    
    
    
    
    
    
    
    
    
    