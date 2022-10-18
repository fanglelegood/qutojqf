#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QDebug>
#include <QClipboard>
#include <QApplication>
#include <QDir>
#include <QtCore>
#include <QDesktopServices>
#include <io.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <sys/locking.h>
#include <share.h>
#include <fcntl.h>
#include "xlsxdocument.h"
#include <QDebug>
#include <QString>
#include <Python.h>



MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setWindowTitle("JQF自动生成");
//    this->setWindowIcon(QIcon(":/test/test.jpg"));

    QAction *action = new QAction(this);
    action->setShortcut(tr("ctrl+v"));
    this->addAction(action);
    connect(action, SIGNAL(triggered()), this, SLOT(autopaste()));
    QAction *action1 = new QAction(this);
    action1->setShortcut(tr("ctrl+c"));
    this->addAction(action1);
    connect(action1, SIGNAL(triggered()), this, SLOT(copySelectFromTable()));

}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked()
{
//     QString nowfile = QDir::currentPath();
     QString strFileOrgcwd="D:\\qutojqf\\clockwiseCAD\\JQF动叶片 顺时针-20210610.dwg";
//     QString strFileOrgcwd=nowfile + "\\clockwiseCAD\\JQF动叶片 顺时针-20210610.dwg";
     QString strFileOrgcwj1="D:\\qutojqf\\clockwiseCAD\\JQF导叶片 顺时针-20210610.dwg";
     QString strFileOrgcwj2="D:\\qutojqf\\clockwiseCAD\\JQF末导叶根 顺时针20210610.dwg";
     QString strFileOrgcwj3="D:\\qutojqf\\clockwiseCAD\\JQF末导叶片 顺时针20210610.dwg";
     QString strFileOrgcwj4="D:\\qutojqf\\clockwiseCAD\\JQF首导叶片 顺时针-20210610.dwg";

     QString strFileOrgccwd="D:\\qutojqf\\CounterclockwiseCAD\\JQF动叶片 逆时针-20210610.dwg";
     QString strFileOrgccwj1="D:\\qutojqf\\CounterclockwiseCAD\\JQF导叶片 逆时针-20210610.dwg";
     QString strFileOrgccwj2="D:\\qutojqf\\CounterclockwiseCAD\\JQF末导叶根 逆时针20210610.dwg";
     QString strFileOrgccwj3="D:\\qutojqf\\CounterclockwiseCAD\\JQF末导叶片 逆时针20210610.dwg";
     QString strFileOrgccwj4="D:\\qutojqf\\CounterclockwiseCAD\\JQF首导叶片 逆时针-20210610.dwg";

//     QFile::copy(strFileOrgcwd, nowfile + "动叶片.dwg");//把fFileOrg复制到strFilePathCopy路径下

    for (int i=0 ; i<ui->lineEdit_2 ->text().toInt(); i++ ) {
        QString strdirection = ui->tableWidget->item(i,18)->text();
        QString strFigureNamed = ui->tableWidget->item(i,9)->text();
        QString strFIleNamed = ui->tableWidget->item(i,8)->text();
        QString strFigureNamej = ui->tableWidget->item(i,2)->text();
        QString strFIleNamej = ui->tableWidget->item(i,1)->text();
//        QString nbexcel = ui->tableWidget->item(i,0)->text(); //取excel中级号可以写入excel中
//        qDebug() << nbexcel  << Qt::endl ;


        if(strdirection == "顺"){
           QFile::copy(strFileOrgcwd, strFileCopyPath + strFIleNamed + ".01.dwg");//把fFileOrg复制到strFilePathCopy路径下
           QFile::copy(strFileOrgcwj1, strFileCopyPath + strFIleNamej + ".01.dwg");
           QFile::copy(strFileOrgcwj2, strFileCopyPath + strFIleNamej + ".04.dwg");
           QFile::copy(strFileOrgcwj3, strFileCopyPath + strFIleNamej + ".03.dwg");
           QFile::copy(strFileOrgcwj4, strFileCopyPath + strFIleNamej + ".02.dwg");
        } else if(strdirection == "逆"){
            QFile::copy(strFileOrgccwd, strFileCopyPath + strFIleNamed + ".01.dwg");//把fFileOrg复制到strFilePathCopy路径下
            QFile::copy(strFileOrgccwj1, strFileCopyPath + strFIleNamej + ".01.dwg");
            QFile::copy(strFileOrgccwj2, strFileCopyPath + strFIleNamej + ".04.dwg");
            QFile::copy(strFileOrgccwj3, strFileCopyPath + strFIleNamej + ".03.dwg");
            QFile::copy(strFileOrgccwj4, strFileCopyPath + strFIleNamej + ".02.dwg");
        }




//        以下输出lsp文件
        // 动叶输出lsp

        //读取python输出的相应txt

        QFile file("D:\\qutojqf\\new\\Bladedata\\"+QString::number(i+1)+"-d.txt" );
        file.open(QIODevice::ReadOnly | QIODevice::Text);
        QTextCodec *codec = QTextCodec::codecForName("GBK");
        QString mbladedatastring=codec->toUnicode(file.readAll());




        QString zwmframemain1 =R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样代号" ")" + strFIleNamed + R"(.01" "Y"))" +
                R"((command "GATTE" "b" "ZwmFrameMain_图样代号" "图样代号" ")" + strFIleNamed + R"(.01" "Y"))" +
            R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样名称" ")" + strFigureNamed +  R"(动叶片" "Y"))" + mbladedatastring ;
//            R"((command "GATTE" "b" "动叶12.5-56.7" "A" ")" + "16" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "AG0" ")" + "14" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "AB0" ")" + "1" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "ALA" ")" + "14" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "AB" ")" + "0.7" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "AB1" ")" + "(1.2)" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "HLA" ")" + "59±0.1" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "HLA0" ")" + "59±0.1" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "HG0" ")" + "66.5±0.2" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "HGA" ")" + "20.5" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "HA" ")" + "93±0.2" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "HGA1" ")" + "27" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "RG" ")" + "2-R0.8" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "DA1" ")" + "(%%C564)" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "RLE1" ")" + "R2" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "RLE3" ")" + "R2" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "RLE2" ")" + "R2" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "RLE4" ")" + "R2" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "DA" ")" + "%%C580" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "CF" ")" + "C0.5" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "AM" ")" + "16" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "H3" ")" + "0" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "DQ" ")" + "%%C0" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "安装角B" ")" + "58.7°" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "L" ")" + "16.91" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "LP1" ")" + "0.05" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "LP2" ")" + "0.05" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "YH" ")" + "5.39±0.1" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "RG0" ")" + "R0.5" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "P" ")" + "14.38±0.05" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "Q" ")" + "10.81±0.05" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "菱形角" ")" + "20°" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "B" ")" + "9.31±0.1" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "YS" ")" + "9.59±0.1" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "YD" ")" + "(4.34)" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "叶型号" ")" + "HS43024.05.06.50(JQF1-16)" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "叶根轮槽图号" ")" + "HS33042.25.01.19" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "PH" ")" + "14.38+0.3" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "QH" ")" + "10.81+0.3" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "ZA" ")" + "154" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "ZA1" ")" + "153" + R"(" "Y"))"
//            R"((command "GATTE" "b" "动叶12.5-56.7" "ZH" ")" + "15" + R"(" "Y"))";
         QString zwmframemainclose =R"((command "._close" "N"))" ;
//            qDebug() << zwmframemain1  << Qt::endl ;


        QFile myfile("D:\\qutojqf\\autojqf.lsp");//创建一个输出文件的文档
        if (myfile.open(QFile::WriteOnly|QFile::Truncate))//注意WriteOnly是往文本中写入的时候用，ReadOnly是在读文本中内容的时候用，Truncate表示将原来文件中的内容清空
            {
                QTextStream out(&myfile);
              out<<zwmframemain1 << Qt::endl;
              out<<zwmframemainclose << Qt::endl;
            }

        //打开对应cad文件自动写入lsp
        QString filePathd = strFileCopyPath + strFIleNamed + ".01.dwg";
        QString local=QString("file:///%1").arg(filePathd);
        QUrl url(local);
        QDesktopServices::openUrl(url);

        //等待动叶片cad 5s 关闭程序阻塞
        QElapsedTimer t;
        t.start();
        while(t.elapsed()<5000)
        QCoreApplication::processEvents();
        while(isFileUsed(filePathd)){
            QElapsedTimer t;
            t.start();
            while(t.elapsed()<1000)
            QCoreApplication::processEvents();
        }


     // 导叶1输出lsp

        //读取python输出的相应txt

        QFile filej1("D:\\qutojqf\\new\\Bladedata\\"+QString::number(i+1)+"-j1.txt" );
        filej1.open(QIODevice::ReadOnly | QIODevice::Text);
        QTextCodec *codecj1 = QTextCodec::codecForName("GBK");
        QString mbladedatastringj1=codecj1->toUnicode(filej1.readAll());
        QString zwmframemainj1 =R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样代号" ")" + strFIleNamej + R"(.01" "Y"))" +
                R"((command "GATTE" "b" "ZwmFrameMain_图样代号" "图样代号" ")" + strFIleNamej + R"(.01" "Y"))" +
            R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样名称" ")" + strFigureNamej +  R"(导叶片" "Y"))" + mbladedatastringj1 ;
        QFile myfilej1("D:\\qutojqf\\autojqf.lsp");//创建一个输出文件的文档
        if (myfilej1.open(QFile::WriteOnly|QFile::Truncate))//注意WriteOnly是往文本中写入的时候用，ReadOnly是在读文本中内容的时候用，Truncate表示将原来文件中的内容清空
            {
                QTextStream out(&myfilej1);
              out<<zwmframemainj1 << Qt::endl;
              out<<zwmframemainclose << Qt::endl;
            }
        //打开对应cad文件自动写入lsp
        QString filePathdj1 = strFileCopyPath + strFIleNamej + ".01.dwg";
        QString localj1=QString("file:///%1").arg(filePathdj1);
        QUrl urlj1(localj1);
        QDesktopServices::openUrl(urlj1);

        //等待动叶片cad 5s 关闭程序阻塞
//        QElapsedTimer t;
        t.start();
        while(t.elapsed()<5000)
        QCoreApplication::processEvents();
        while(isFileUsed(filePathd)){
            QElapsedTimer t;
            t.start();
            while(t.elapsed()<1000)
            QCoreApplication::processEvents();
        }

    // 导叶2输出lsp

       //读取python输出的相应txt

       QFile filej2("D:\\qutojqf\\new\\Bladedata\\"+QString::number(i+1)+"-j2.txt" );
       filej2.open(QIODevice::ReadOnly | QIODevice::Text);
       QTextCodec *codecj2 = QTextCodec::codecForName("GBK");
       QString mbladedatastringj2=codecj2->toUnicode(filej2.readAll());
       QString zwmframemainj2 =R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样代号" ")" + strFIleNamej + R"(.02" "Y"))" +
               R"((command "GATTE" "b" "ZwmFrameMain_图样代号" "图样代号" ")" + strFIleNamej + R"(.02" "Y"))" +
           R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样名称" ")" + strFigureNamej +  R"(首导叶片" "Y"))" + mbladedatastringj2 ;
       QFile myfilej2("D:\\qutojqf\\autojqf.lsp");//创建一个输出文件的文档
       if (myfilej2.open(QFile::WriteOnly|QFile::Truncate))//注意WriteOnly是往文本中写入的时候用，ReadOnly是在读文本中内容的时候用，Truncate表示将原来文件中的内容清空
           {
               QTextStream out(&myfilej2);
             out<<zwmframemainj2 << Qt::endl;
             out<<zwmframemainclose << Qt::endl;
           }
       //打开对应cad文件自动写入lsp
       QString filePathdj2 = strFileCopyPath + strFIleNamej + ".02.dwg";
       QString localj2=QString("file:///%1").arg(filePathdj2);
       QUrl urlj2(localj2);
       QDesktopServices::openUrl(urlj2);

       //等待动叶片cad 5s 关闭程序阻塞
//        QElapsedTimer t;
       t.start();
       while(t.elapsed()<5000)
       QCoreApplication::processEvents();
       while(isFileUsed(filePathd)){
           QElapsedTimer t;
           t.start();
           while(t.elapsed()<1000)
           QCoreApplication::processEvents();
       }

       // 导叶3输出lsp

          //读取python输出的相应txt

          QFile filej3("D:\\qutojqf\\new\\Bladedata\\"+QString::number(i+1)+"-j3.txt" );
          filej3.open(QIODevice::ReadOnly | QIODevice::Text);
          QTextCodec *codecj3 = QTextCodec::codecForName("GBK");
          QString mbladedatastringj3=codecj3->toUnicode(filej3.readAll());
          QString zwmframemainj3 =R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样代号" ")" + strFIleNamej + R"(.03" "Y"))" +
                  R"((command "GATTE" "b" "ZwmFrameMain_图样代号" "图样代号" ")" + strFIleNamej + R"(.03" "Y"))" +
              R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样名称" ")" + strFigureNamej +  R"(末导叶片" "Y"))" + mbladedatastringj3 ;
          QFile myfilej3("D:\\qutojqf\\autojqf.lsp");//创建一个输出文件的文档
          if (myfilej3.open(QFile::WriteOnly|QFile::Truncate))//注意WriteOnly是往文本中写入的时候用，ReadOnly是在读文本中内容的时候用，Truncate表示将原来文件中的内容清空
              {
                  QTextStream out(&myfilej3);
                out<<zwmframemainj3 << Qt::endl;
                out<<zwmframemainclose << Qt::endl;
              }
          //打开对应cad文件自动写入lsp
          QString filePathdj3 = strFileCopyPath + strFIleNamej + ".03.dwg";
          QString localj3=QString("file:///%1").arg(filePathdj3);
          QUrl urlj3(localj3);
          QDesktopServices::openUrl(urlj3);

          //等待动叶片cad 5s 关闭程序阻塞
   //        QElapsedTimer t;
          t.start();
          while(t.elapsed()<5000)
          QCoreApplication::processEvents();
          while(isFileUsed(filePathd)){
              QElapsedTimer t;
              t.start();
              while(t.elapsed()<1000)
              QCoreApplication::processEvents();
          }

          // 导叶4输出lsp

             //读取python输出的相应txt

             QFile filej4("D:\\qutojqf\\new\\Bladedata\\"+QString::number(i+1)+"-j4.txt" );
             filej4.open(QIODevice::ReadOnly | QIODevice::Text);
             QTextCodec *codecj4 = QTextCodec::codecForName("GBK");
             QString mbladedatastringj4=codecj4->toUnicode(filej4.readAll());
             QString zwmframemainj4 =R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样代号" ")" + strFIleNamej + R"(.04" "Y"))" +
                     R"((command "GATTE" "b" "ZwmFrameMain_图样代号" "图样代号" ")" + strFIleNamej + R"(.04" "Y"))" +
                 R"((command "GATTE" "b" "ZwmFrameMain_中文图纸标题栏" "图样名称" ")" + strFigureNamej +  R"(末导叶根" "Y"))" + mbladedatastringj4 ;
             QFile myfilej4("D:\\qutojqf\\autojqf.lsp");//创建一个输出文件的文档
             if (myfilej4.open(QFile::WriteOnly|QFile::Truncate))//注意WriteOnly是往文本中写入的时候用，ReadOnly是在读文本中内容的时候用，Truncate表示将原来文件中的内容清空
                 {
                     QTextStream out(&myfilej4);
                   out<<zwmframemainj4 << Qt::endl;
                   out<<zwmframemainclose << Qt::endl;
                 }
             //打开对应cad文件自动写入lsp
             QString filePathdj4 = strFileCopyPath + strFIleNamej + ".04.dwg";
             QString localj4=QString("file:///%1").arg(filePathdj4);
             QUrl urlj4(localj4);
             QDesktopServices::openUrl(urlj4);

             //等待动叶片cad 5s 关闭程序阻塞
      //        QElapsedTimer t;
             t.start();
             while(t.elapsed()<5000)
             QCoreApplication::processEvents();
             while(isFileUsed(filePathd)){
                 QElapsedTimer t;
                 t.start();
                 while(t.elapsed()<1000)
                 QCoreApplication::processEvents();
             }
    }


    QFile myfilej5("D:\\qutojqf\\autojqf.lsp");//创建一个输出文件的文档
    if (myfilej5.open(QFile::WriteOnly|QFile::Truncate))//注意WriteOnly是往文本中写入的时候用，ReadOnly是在读文本中内容的时候用，Truncate表示将原来文件中的内容清空
        {

        }
    QMessageBox::information(this,"提示信息","完成生成请至软件目录new文件夹查阅");



    // QString strFIleName="1234";//新的名字




    // QFile::copy(strFileOrg, strFileCopyPath + strFIleName + ".dwg");//把fFileOrg复制到strFilePathCopy路径下
//    qDebug() << strFIleName  << endl;


    //当然，每一个路径和名字都是可以通过读取得来，比如QString strFIleName=pFile->nFileID；
}


// 复制选中内容
void MainWindow::copySelectFromTable()
{
    QList<QTableWidgetSelectionRange> sRangeList = ui->tableWidget->selectedRanges();
    for(const auto &p : qAsConst(sRangeList)) {


        QString str;
        for (auto i = p.topRow(); i <= p.bottomRow(); i++) {
            QString rowStr;
            for (auto j = p.leftColumn(); j <= p.rightColumn(); j++) {
                QTableWidgetItem* item =  ui->tableWidget->item(i, j);
                if(item != nullptr) {
                    if(j == p.leftColumn())
                        rowStr = item->text() + "\t";
                    else if (j == p.rightColumn())
                        rowStr = rowStr + item->text() + "\n";
                    else
                        rowStr = rowStr + item->text() + "\t";
                }
                else {
                    break;
                }
            }
            str += rowStr;
        }
        QClipboard *pClip= QApplication::clipboard();
        pClip->setText(str);
    }
}

// 粘贴，从选中的第一个单元格开始


void MainWindow::autopaste()
{
    QList<QTableWidgetSelectionRange> sRangeList = ui->tableWidget->selectedRanges();
    int allRow = ui->tableWidget->rowCount();
    int allCol = ui->tableWidget->columnCount();
    for(const auto &p : qAsConst(sRangeList)) {
        QClipboard *pClip= QApplication::clipboard();
        QString str = pClip->text();
        int ColCnt = ui->tableWidget->columnCount();
        QList<QString> RowStr = str.split("\n");
        int copyAreaAllRow = RowStr.size();
        qDebug()<<"copyAreaAllRow"<<copyAreaAllRow;
        int x = p.topRow();
        int rightIndex = p.rightColumn();
        int surplusRow = allRow - x;
        int surplusCol = allCol - rightIndex;
        if((RowStr.size() -1) > surplusRow)//如果复制的行数大于剩余的行数，去除掉多余的赋值内容
        {
            int len = RowStr.size();
            for(int i=len -2;i>=surplusRow;i--)
            {
                RowStr.removeAt(i);
            }
        }


        for(const auto &Row : qAsConst(RowStr)) {
            if(!Row.isEmpty()) {
                QList<QString> ColStr = Row.split("\t");//赋值的列数
                if(ColStr.size() > surplusCol)//如果复制的列数大于剩余的列数，去除掉多余的赋值内容
                {
                    int len = ColStr.size();
                    for(int i=len -1;i>=surplusCol;i--)
                    {
                        ColStr.removeAt(i);
                    }
                }
                int y = p.leftColumn();
                for(const auto &Col : qAsConst(ColStr)) {
                    QTableWidgetItem* item = ui->tableWidget->item(x, y);
                    if(item == nullptr)
                    {
                        ui->tableWidget->setItem(x, y, new QTableWidgetItem(Col));
                        ui->tableWidget->item(x,y)->setForeground(QBrush(QColor(255,0,0)));
                    }
                    else
                    {
                        ui->tableWidget->item(x, y)->setText(Col);
                        ui->tableWidget->item(x,y)->setForeground(QBrush(QColor(255,0,0)));
                    }
                    if(y + 1 == ColCnt)
                        break;
                    ++y;
                }
                ++x;
            }
        }
    }
}

bool MainWindow::isFileUsed(QString fpath)
{
    bool isUse = false;

    QString wpath = QDir::toNativeSeparators(fpath);

    QByteArray qbyteArr;
    qbyteArr.append(wpath);
    const char *c2 = qbyteArr.data();

    int fh = _sopen(c2, _O_RDWR, _SH_DENYRW,_S_IREAD | _S_IWRITE );
    if(-1 == fh)
    {
        isUse = true;
    }
    else
    {
        _close(fh);
    }

    return isUse;
}





//void MainWindow::on_pushButton_2_clicked()
//{
//    diad.show();
//}


void MainWindow::on_pushButton_3_clicked()
{
    QString strFiletem="D:\\qutojqf\\autoJQF输入模板.xlsx";
    if(ui->lineEdit->text().isEmpty()){
        QMessageBox::information(this,"提示信息","请输入机组号");
        return;
    }
    QFile::copy(strFiletem, strFileCopyPath + ui->lineEdit->text() + "Flowdata.xlsx");//把fFileOrg复制到strFilePathCopy路径下
    QString filePathtem = strFileCopyPath + ui->lineEdit->text() + "Flowdata.xlsx";
    QString local=QString("file:///%1").arg(filePathtem);
    QUrl url(local);
    QDesktopServices::openUrl(url);

}


void MainWindow::on_pushButton_4_clicked()
{
    QString filePathtem = strFileCopyPath + ui->lineEdit->text() + "Flowdata.xlsx";
//    qDebug() << filePathtem  << Qt::endl;
//    qDebug() << isFileUsed(filePathtem)  << Qt::endl;

    if(isFileUsed(filePathtem)){
        QMessageBox::information(this,"提示信息","请先关闭并保存输入excel");
        return;
    }

    QXlsx::Document xlsx(filePathtem);

    for (int i =0; i<ui->tableWidget->rowCount();i++){
        for (int j =0; j<ui->tableWidget->columnCount();j++){
            ui->tableWidget->setItem(i,j,new QTableWidgetItem(""));
            ui->tableWidget->item(i,j)->setTextAlignment(Qt::AlignCenter);
            ui->tableWidget->item(i,j)->setText(xlsx.read(i+3,j+1).toString());
        }
    }


    xlsx.save();
}


void MainWindow::on_pushButton_5_clicked()
{

}


void MainWindow::on_pushButton_7_clicked()
{
    QString strFiletem="D:\\qutojqf\\JQF叶片绘图表-20220113.xlsx";
    if(ui->lineEdit->text().isEmpty()){
        QMessageBox::information(this,"提示信息","请输入机组号");
        return;
    }
    QFile::copy(strFiletem, strFileCopyPath + ui->lineEdit->text() + "JQFbladedrawingtable.xlsx");//把fFileOrg复制到strFilePathCopy路径下
    QString filePathtem = strFileCopyPath + ui->lineEdit->text() + "JQFbladedrawingtable.xlsx";
    QString local=QString("file:///%1").arg(filePathtem);
    QUrl url(local);
    QDesktopServices::openUrl(url);
}


void MainWindow::on_pushButton_6_clicked()
{
    if(ui->lineEdit->text().isEmpty()){
        QMessageBox::information(this,"提示信息","请输入机组号");
        return;
    }
    if(ui->lineEdit_2 ->text().isEmpty()){
        QMessageBox::information(this,"提示信息","请输入总共需要计算的级数");
        return;
    }

    QString strFileUnit = strFileCopyPath+ ui->lineEdit->text() + "JQFbladedrawingtable.xlsx" ;
    QString series = ui->lineEdit_2 ->text();

    //调用python读取excel生成lsp文件
        Py_SetPythonHome((wchar_t *)(L"./Python38"));
        Py_Initialize();
        PyObject * pModule = NULL;
        PyObject * pFunc = NULL;
        PyObject *pDict = NULL;


        pModule = PyImport_ImportModule("readexcel");
        pDict = PyModule_GetDict(pModule);
        pFunc = PyDict_GetItemString(pDict, "GenerateTxt");

        PyObject *pArgs = PyTuple_New(2);//函数调用的参数传递均是以元组的形式打包的,2表示参数个数
        PyTuple_SetItem(pArgs, 0, Py_BuildValue("i", series.toInt()));//0--序号,i表示创建int型变量
        PyTuple_SetItem(pArgs, 1, Py_BuildValue("s", strFileUnit.toStdString().c_str()));

        PyObject_CallObject(pFunc, pArgs);//调用函数，完成传递

        Py_Finalize();

}

