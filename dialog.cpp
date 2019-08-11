#pragma execution_character_set("utf-8")
#include "dialog.h"
#include "ui_dialog.h"
#include <QFileDialog>
#include <ActiveQt/QAxObject>
#include <QMessageBox>
#include <QDebug>
#include "exceloperator.h"
#include <iostream>
#include <QJsonDocument>
#include <QJsonParseError> 
#include <QJsonArray>
#include <QJsonObject>
#include <QPair>
Dialog::Dialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Dialog)
{
    ui->setupUi(this);
    pModel = new QStandardItemModel(this);
    pSelection = new QItemSelectionModel(pModel);
    ui->tableView->setModel(pModel);
    ui->tableView->setSelectionModel(pSelection);
    pModel->setHorizontalHeaderLabels(QStringList() << "孔名称" << "X" << "Y" << "Z" << "+TOL" << "-TOL" << "R" << "+TOL" << "-TOL");
    ui->tableView->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    //ui->tableView->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    //tableView->horizontalHeader()->setMinimumSectionSize(100);
}

Dialog::~Dialog()
{
    delete ui;
}
void Dialog::on_btnOpenExcel_clicked(){
    QString excelPath = QFileDialog::getOpenFileName(this,"请打开Excel文件",QDir::currentPath(),"Excel文件(*.xlsx)");
    if(excelPath.isEmpty()){
        QMessageBox::information(this,"信息","请打开Excel文件");
        return;
    }
    pModel->clear();
    pModel->setHorizontalHeaderLabels(QStringList() << "孔名称" << "X" << "Y" << "Z" << "+TOL" << "-TOL" << "R" << "+TOL" << "-TOL");
    
    qDebug()<<excelPath;
    GetFileContent(excelPath);
}
void Dialog::GetFileContent(QString filename)
{
	if (filename.isEmpty()) {
		return ;
	}
    ExcelOperator readExcel;
    bool bl = readExcel.open(filename, false);
    if (!bl) {
        readExcel.close();
        QMessageBox::information(this, "信息", "请打开Excel文件错误");
        return;
    }
    QAxObject* pWorkSheet = readExcel.getSheet(1);
    if (pWorkSheet == NULL) {
        readExcel.close();
        QMessageBox::information(this, "信息", "Excel没有找到第一个Sheet表");
        return;
    }
    int row_count = readExcel.getRowsCount(pWorkSheet);
    int col_count = readExcel.getColumnsCount(pWorkSheet);
    qDebug() << "Excel row = " << row_count << ", cols = " << col_count;

    QJsonParseError error;
    QJsonDocument jsonDoc1;
    QJsonDocument jsonDoc2;
    QJsonArray jsonArr1;
    QJsonArray jsonArr2;

    for (int i = 2; i < row_count + 1; ++i) {
         QString name = readExcel.getCell(pWorkSheet, i, 1);
         if (name.isEmpty()) {
             break;
         }
         bool ok;
         double X = readExcel.getCell(pWorkSheet, i, 2).toDouble(&ok);
         if (!ok) {
             X = NAN;
         }
         double Y = readExcel.getCell(pWorkSheet, i, 3).toDouble(&ok);
         if (!ok) {
             Y = NAN;
         }
         double Z = readExcel.getCell(pWorkSheet, i, 4).toDouble(&ok);
         if (!ok) {
             Z = NAN;
         }
         double TOL1 = readExcel.getCell(pWorkSheet, i, 5).toDouble(&ok);
         if (!ok) {
             TOL1 = NAN;
         }
         double TOL2 = readExcel.getCell(pWorkSheet, i, 6).toDouble(&ok);
         if (!ok) {
             TOL2 = NAN;
         }
         double R = readExcel.getCell(pWorkSheet, i, 7).toDouble(&ok);
         if (!ok) {
             X = NAN;
         }
         double RTOL1 = readExcel.getCell(pWorkSheet, i, 8).toDouble(&ok);
         if (!ok) {
             RTOL1 = NAN;
         }
         double RTOL2 = readExcel.getCell(pWorkSheet, i, 9).toDouble(&ok);
         if (!ok) {
             RTOL2 = NAN;
         }
         int row = pModel->rowCount();
         pModel->insertRow(row);
         pModel->setItem(row, 0, new QStandardItem(name));
         pModel->setItem(row, 1, new QStandardItem(QString("%1").arg(X)));
         pModel->setItem(row, 2, new QStandardItem(QString("%1").arg(Y)));
         pModel->setItem(row, 3, new QStandardItem(QString("%1").arg(Z)));
         pModel->setItem(row, 4, new QStandardItem(QString("%1").arg(TOL1)));
         pModel->setItem(row, 5, new QStandardItem(QString("%1").arg(TOL2)));
         pModel->setItem(row, 6, new QStandardItem(QString("%1").arg(R)));
         pModel->setItem(row, 7, new QStandardItem(QString("%1").arg(RTOL1)));
         pModel->setItem(row, 8, new QStandardItem(QString("%1").arg(RTOL2)));

        QJsonObject obj1;	
        QJsonObject obj2;
        obj1.insert("name", name);
        obj1.insert("index", i - 2);
        obj1.insert("x", X);
        obj1.insert("y", Y);
        obj1.insert("z", Z);
        obj1.insert("TOL+", TOL1);
        obj1.insert("TOL-", TOL2);
        obj1.insert("R", R);
        obj1.insert("RTOL+", RTOL1);
        obj1.insert("RTOL-", RTOL2);
        jsonArr1.append(obj1);

        obj2.insert("name", name);
        obj2.insert("index", i - 2);
        obj2.insert("x", X);
        obj2.insert("y", Y);
        obj2.insert("z", Z);
        obj2.insert("R", R);
        jsonArr2.append(obj2);
    }
    readExcel.close();
    ui->tableView->resizeColumnsToContents();
    jsonDoc1.setArray(jsonArr1);
    jsonDoc2.setArray(jsonArr2);

    QFileInfo info(filename);
    qDebug() << info.absolutePath();
    qDebug() << info.baseName();
    QString str = info.absolutePath() + QDir::separator() + info.baseName() + "_1.json";
    QString str2 = info.absolutePath() + QDir::separator() + info.baseName() + "_2.json";
    qDebug() << str;

    QFile fd(str);
    bl = fd.open(QIODevice::WriteOnly | QIODevice::Text);
    if (!bl) {
    	QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
    	return;
    }
    //qDebug()<< jsonDoc1.toJson();
    fd.write(jsonDoc1.toJson());
    fd.close();
    fd.setFileName(str2);

    bl = fd.open(QIODevice::WriteOnly | QIODevice::Text);
    if (!bl) {
    	QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
    	return;
    }
    fd.write(jsonDoc2.toJson());
    fd.close();
    QMessageBox::information(this, "信息", QString("Excel数据读入完成"), QMessageBox::Ok);
 //   QAxObject* excel = new QAxObject();
 //   if (!excel->setControl("Excel.Application"))//判断是否成功连接excel文件
 //   {
 //       delete excel;
 //       excel = NULL;
 //       //CoUninitialize();
 //       return -2;
 //   }
 //   excel->setProperty("Visible", false);
 //   QAxObject* work_books = excel->querySubObject("WorkBooks");
 //   work_books->dynamicCall("Open (const QString&)", filename);
 //   if (work_books == NULL)
 //   {
 //       excel->dynamicCall("Quit(void)"); //退出
 //       delete excel;
 //       excel = NULL;
 //       //CoUninitialize();
 //       return -1;
 //   }
 //   QAxObject* work_book = excel->querySubObject("ActiveworkBook");
 //   if (work_book == NULL)
 //   {
 //       work_books->dynamicCall("Close(Boolean)", false); //关闭文件
 //       excel->dynamicCall("Quit(void)"); //退出
 //       delete work_books;
 //       work_books = NULL;
 //       delete excel;
 //       excel = NULL;
 //       //CoUninitialize();
 //       return -1;
 //   }
 //   QAxObject* work_sheets = work_book->querySubObject("Worksheets(int)", 1);  //Sheets也可换用WorkSheets
 //   QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
 //   QAxObject *used_range = work_sheet->querySubObject("UsedRange");
 //   QAxObject *rows = used_range->querySubObject("Rows");
 //   QAxObject *columns = used_range->querySubObject("Columns");
 //   int row_count = rows->property("Count").toInt();  //获取行数
 //   int col_count = columns->property("Count").toInt();  //获取列数


	
	//

 //   for (int i = 1; i < row_count; i++)
 //   {
	//	QJsonObject obj1;	
	//	QJsonObject obj2;
	//	obj1.insert("name", work_sheet->querySubObject("Cells(int,int)", i + 1, 1)->dynamicCall("Value").toString());
	//	obj1.insert("index", i - 1);
	//	int type = work_sheet->querySubObject("Cells(int,int)", i + 1, 2)->dynamicCall("Value").toInt();
	//	obj1.insert("type", type);
	//	//QJsonObject objMin;
	//	QJsonObject objMinVal;
	//	//QJsonObject objNormal;
	//	QJsonObject objNormalVal;
	//	//QJsonObject objMax;
	//	QJsonObject objMaxVal;
	//	switch (type)
	//	{
	//	case 0:
	//	{
	//		objMinVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 3)->dynamicCall("Value").toDouble());
	//		objMinVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 4)->dynamicCall("Value").toDouble());
	//		objMinVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 5)->dynamicCall("Value").toDouble());

	//		objNormalVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
	//		objNormalVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 7)->dynamicCall("Value").toDouble());
	//		objNormalVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 8)->dynamicCall("Value").toDouble());

	//		objMaxVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 9)->dynamicCall("Value").toDouble());
	//		objMaxVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 10)->dynamicCall("Value").toDouble());
	//		objMaxVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 11)->dynamicCall("Value").toDouble());
	//	}
	//	break;
	//	case 1:
	//		objMinVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 3)->dynamicCall("Value").toDouble());
	//		objNormalVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
	//		objMaxVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 9)->dynamicCall("Value").toDouble());
	//		break;
	//	default:
	//		QMessageBox::information(this, "类型错误", QString("未知类型：{0}").arg(type), QMessageBox::Ok);
	//		return false;
	//	}
	//	obj1.insert("norm_value", objNormalVal);
	//	obj1.insert("min_value", objMinVal);
	//	obj1.insert("max_value", objMaxVal);
	//	jsonArr1.append(obj1);

	//	obj2.insert("index", i - 1);
	//	obj2.insert("type", type);
	//	QJsonObject objVal;
	//	switch (type)
	//	{
	//	case 0:
	//		objVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
	//		objVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 7)->dynamicCall("Value").toDouble());
	//		objVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 8)->dynamicCall("Value").toDouble());
	//		break;
	//	case 1:
	//		objNormalVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
	//		break;
	//	}
	//	obj2.insert("value", objVal);
	//	jsonArr2.append(obj2);
 //   }
	//jsonDoc1.setArray(jsonArr1);
	//jsonDoc2.setArray(jsonArr2);

 //   work_books->dynamicCall("Close(Boolean)", false);
 //   excel->dynamicCall("Quit(void)"); //退出
 //   delete work_books;
 //   work_books = NULL;
 //   delete excel;
 //   excel = NULL;
    //CoUninitialize();
	//QFileInfo info(filename);
	//qDebug() << info.absolutePath();
	//qDebug() << info.baseName();
	//QString str = info.absolutePath() + QDir::separator() + info.baseName() + "_1.json";
	//QString str2 = info.absolutePath() + QDir::separator() + info.baseName() + "_2.json";
	//qDebug() << str;
	//QFile fd(str);
	//bool bl = fd.open(QIODevice::WriteOnly | QIODevice::Text);
	//if (!bl) {
	//	QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
	//	return 0;
	//}
	////qDebug()<< jsonDoc1.toJson();
	//fd.write(jsonDoc1.toJson());
	//fd.close();
	//fd.setFileName(str2);
	//bl = fd.open(QIODevice::WriteOnly | QIODevice::Text);
	//if (!bl) {
	//	QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
	//	return 0;
	//}
	//fd.write(jsonDoc2.toJson());
	//fd.close();
	//QMessageBox::information(this, "信息", QString("文件写入成功"), QMessageBox::Ok);
 //   
    return;
}