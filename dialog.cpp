#pragma execution_character_set("utf-8")
#include "dialog.h"
#include "ui_dialog.h"
#include <QFileDialog>
#include <ActiveQt/QAxObject>
#include <QMessageBox>
#include <QDebug>
#include "exceloperator.h"
#include <iostream>
#include <objbase.h>
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
	/*QJsonDocument doc;
	QJsonArray jsonArr;
	
	QJsonObject obj1;
	
	obj1.insert("name", "name1");
	obj1.insert("index", 0);
	QJsonObject obj2;
	
	obj2.insert("name", "name1");
	obj2.insert("index", 0);
	jsonArr.append(obj1);
	jsonArr.append(obj2);
	doc.setArray(jsonArr);
	qDebug() << doc.toJson();*/
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
    qDebug()<<excelPath;
	/*QFileInfo info(excelPath);
	qDebug()<< info.absolutePath();
	qDebug() << info.baseName();*/
    GetFileContent(excelPath);
//    std::cout<<"aaa"<<std::flush<<std::end;
    //excelPath.toLocal8Bit()
//    ExcelOperator* excel = new ExcelOperator(this);
//    bool bl = excel->open(excelPath);

//    qDebug()<<"excel->open = "<<bl;

//    QAxObject excel("Excel.Application");
//    excel.setProperty("Visible", false);
//    QAxObject *work_books = excel.querySubObject("WorkBooks");
//    work_books->dynamicCall("Open (const QString&)", excelPath);
//    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
//    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
//    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目

//    if(sheet_count > 0)
//    {
//        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
//        QAxObject *used_range = work_sheet->querySubObject("UsedRange");
//        QAxObject *rows = used_range->querySubObject("Rows");
//        //QAxObject *columns = used_range->querySubObject("Columns");
//        int row_start = used_range->property("Row").toInt();  //获取起始行:1
//        //int column_start = used_range->property("Column").toInt();  //获取起始列
//        int row_count = rows->property("Count").toInt();  //获取行数
//        //int column_count = columns->property("Count").toInt();  //获取列数

//        QString StudentName;
//        for(int i=row_start+4; i<=row_count;i++)
//        {
//            QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", i, 3);
//            StudentName = cell->dynamicCall("Value2()").toString();//获取单元格内容
//// 			cell = work_sheet->querySubObject("Cells(int,int)", i, 3);
//// 			StudentNum[i-1] = cell->dynamicCall("Value2()").toString();//获取(i,3)

////            stuNames.push_back(StudentName.toStdString());
//        }
//    }

//    work_books->dynamicCall("Close()");//关闭工作簿
//    excel.dynamicCall("Quit()");//关闭excel
}
int Dialog::GetFileContent(QString filename)
{
	if (filename.isEmpty()) {
		return 0;
	}
    QString Text;
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    //fileContent.clear();
    QAxObject* excel = new QAxObject();
    if (!excel->setControl("Excel.Application"))//判断是否成功连接excel文件
    {
        delete excel;
        excel = NULL;
        CoUninitialize();
        return -2;
    }
    excel->setProperty("Visible", false);
    QAxObject* work_books = excel->querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)", filename);
    if (work_books == NULL)
    {
        excel->dynamicCall("Quit(void)"); //退出
        delete excel;
        excel = NULL;
        CoUninitialize();
        return -1;
    }
    QAxObject* work_book = excel->querySubObject("ActiveworkBook");
    if (work_book == NULL)
    {
        work_books->dynamicCall("Close(Boolean)", false); //关闭文件
        excel->dynamicCall("Quit(void)"); //退出
        delete work_books;
        work_books = NULL;
        delete excel;
        excel = NULL;
        CoUninitialize();
        return -1;
    }
    QAxObject* work_sheets = work_book->querySubObject("Worksheets(int)", 1);  //Sheets也可换用WorkSheets
    QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
    QAxObject *used_range = work_sheet->querySubObject("UsedRange");
    QAxObject *rows = used_range->querySubObject("Rows");
    QAxObject *columns = used_range->querySubObject("Columns");
    int row_count = rows->property("Count").toInt();  //获取行数
    int col_count = columns->property("Count").toInt();  //获取列数


	QJsonParseError error;
	QJsonDocument jsonDoc1;
	QJsonDocument jsonDoc2;
	QJsonArray jsonArr1;
	QJsonArray jsonArr2;
	

    for (int i = 1; i < row_count; i++)
    {
		QJsonObject obj1;	
		QJsonObject obj2;
		obj1.insert("name", work_sheet->querySubObject("Cells(int,int)", i + 1, 1)->dynamicCall("Value").toString());
		obj1.insert("index", i - 1);
		int type = work_sheet->querySubObject("Cells(int,int)", i + 1, 2)->dynamicCall("Value").toInt();
		obj1.insert("type", type);
		//QJsonObject objMin;
		QJsonObject objMinVal;
		//QJsonObject objNormal;
		QJsonObject objNormalVal;
		//QJsonObject objMax;
		QJsonObject objMaxVal;
		switch (type)
		{
		case 0:
		{
			objMinVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 3)->dynamicCall("Value").toDouble());
			objMinVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 4)->dynamicCall("Value").toDouble());
			objMinVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 5)->dynamicCall("Value").toDouble());

			objNormalVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
			objNormalVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 7)->dynamicCall("Value").toDouble());
			objNormalVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 8)->dynamicCall("Value").toDouble());

			objMaxVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 9)->dynamicCall("Value").toDouble());
			objMaxVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 10)->dynamicCall("Value").toDouble());
			objMaxVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 11)->dynamicCall("Value").toDouble());
		}
		break;
		case 1:
			objMinVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 3)->dynamicCall("Value").toDouble());
			objNormalVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
			objMaxVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 9)->dynamicCall("Value").toDouble());
			break;
		default:
			QMessageBox::information(this, "类型错误", QString("未知类型：{0}").arg(type), QMessageBox::Ok);
			return false;
		}
		obj1.insert("norm_value", objNormalVal);
		obj1.insert("min_value", objMinVal);
		obj1.insert("max_value", objMaxVal);
		jsonArr1.append(obj1);

		obj2.insert("index", i - 1);
		obj2.insert("type", type);
		QJsonObject objVal;
		switch (type)
		{
		case 0:
			objVal.insert("x", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
			objVal.insert("y", work_sheet->querySubObject("Cells(int,int)", i + 1, 7)->dynamicCall("Value").toDouble());
			objVal.insert("z", work_sheet->querySubObject("Cells(int,int)", i + 1, 8)->dynamicCall("Value").toDouble());
			break;
		case 1:
			objNormalVal.insert("value", work_sheet->querySubObject("Cells(int,int)", i + 1, 6)->dynamicCall("Value").toDouble());
			break;
		}
		obj2.insert("value", objVal);
		jsonArr2.append(obj2);
        /*for (int j = 0; j < col_count; j++)
        {
			obj1.
            Text = work_sheet->querySubObject("Cells(int,int)",i+1,j+1)->dynamicCall("Value").toString();
        }*/
    }
	jsonDoc1.setArray(jsonArr1);
	jsonDoc2.setArray(jsonArr2);

    work_books->dynamicCall("Close(Boolean)", false);
    excel->dynamicCall("Quit(void)"); //退出
    delete work_books;
    work_books = NULL;
    delete excel;
    excel = NULL;
    CoUninitialize();
	QFileInfo info(filename);
	qDebug() << info.absolutePath();
	qDebug() << info.baseName();
	QString str = info.absolutePath() + QDir::separator() + info.baseName() + "_1.json";
	QString str2 = info.absolutePath() + QDir::separator() + info.baseName() + "_2.json";
	qDebug() << str;
	QFile fd(str);
	bool bl = fd.open(QIODevice::WriteOnly | QIODevice::Text);
	if (!bl) {
		QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
		return 0;
	}
	//qDebug()<< jsonDoc1.toJson();
	fd.write(jsonDoc1.toJson());
	fd.close();
	fd.setFileName(str2);
	bl = fd.open(QIODevice::WriteOnly | QIODevice::Text);
	if (!bl) {
		QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
		return 0;
	}
	//qDebug() << jsonDoc2.toJson();
	fd.write(jsonDoc2.toJson());
	fd.close();
	QMessageBox::information(this, "信息", QString("文件写入成功"), QMessageBox::Ok);
	/*FILE* fd = fopen(path, "wt");
	if (!fd) {
		QMessageBox::information(this, "错误", QString("文件写入"), QMessageBox::Ok);
		return 0;
	}
	fwrite()*/
    return 1;
}
/*
bool Dialog::xlsReader(QString excelPath,vector<string> &stuNames)
{
    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false);
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)", excelPath);
    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目

    if(sheet_count > 0)
    {
        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
        QAxObject *used_range = work_sheet->querySubObject("UsedRange");
        QAxObject *rows = used_range->querySubObject("Rows");
        //QAxObject *columns = used_range->querySubObject("Columns");
        int row_start = used_range->property("Row").toInt();  //获取起始行:1
        //int column_start = used_range->property("Column").toInt();  //获取起始列
        int row_count = rows->property("Count").toInt();  //获取行数
        //int column_count = columns->property("Count").toInt();  //获取列数

        QString StudentName;
        for(int i=row_start+4; i<=row_count;i++)
        {
            QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", i, 3);
            StudentName = cell->dynamicCall("Value2()").toString();//获取单元格内容
// 			cell = work_sheet->querySubObject("Cells(int,int)", i, 3);
// 			StudentNum[i-1] = cell->dynamicCall("Value2()").toString();//获取(i,3)

            stuNames.push_back(StudentName.toStdString());
        }
    }

    work_books->dynamicCall("Close()");//关闭工作簿
    excel.dynamicCall("Quit()");//关闭excel

    return true;
}
*/