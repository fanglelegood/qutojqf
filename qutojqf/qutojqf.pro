QT       += core gui

# 使用qtxlsx源代码
include(qtxlsx/src/xlsx/qtxlsx.pri)

#QT += xlsx


greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++11

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    dabledated.cpp \
    dialog.cpp \
    main.cpp \
    mainwindow.cpp

HEADERS += \
    dabledated.h \
    dialog.h \
    mainwindow.h

FORMS += \
    dialog.ui \
    mainwindow.ui

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

INCLUDEPATH += -I C:\Users\owner\AppData\Local\Programs\Python\Python38\include
LIBS += -LC:\Users\owner\AppData\Local\Programs\Python\Python38\libs\ -lpython38

DISTFILES += \
    readexcel.py

msvc:QMAKE_CXXFLAGS += -execution-charset:utf-8
msvc:QMAKE_CXXFLAGS += -source-charset:utf-8

RC_ICONS = autocad.ico
