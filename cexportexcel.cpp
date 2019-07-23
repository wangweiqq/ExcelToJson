#pragma execution_character_set("utf-8")
#include "cexportexcel.h"

CExportExcel::CExportExcel(const QAbstractItemView& v,QString filePath,QString sheetName)
    : QThread(nullptr),view(v),mFilePath(filePath),mSheetName(sheetName)
{
    excel = new ExcelOperator(this);
    mIsStop = false;
}
CExportExcel::~CExportExcel(){
    if(this->isRunning()){
//        this->terminate();
        mIsStop = true;
        qDebug()<<"~CExportExcel "<<this->isRunning();
        this->wait();
        excel->close();
        qDebug()<<"CExportExcel 退出";
    }
}
void CExportExcel::run(){
    excel->open(mFilePath);
    if(mIsStop){
        return;
    }
    QAxObject* pSheet = excel->addSheet(mSheetName);
    if(mIsStop){
        return;
    }
    QAbstractItemModel* pModel = view.model();
    int rows = pModel->rowCount();
    int cols = pModel->columnCount();
    emit SetLoadingInfo("正在处理,Excel表头信息");
    //表头
    for(int j = 0;j<cols;++j){
        if(mIsStop){
            return;
        }
        QString value = pModel->headerData(j,Qt::Horizontal,Qt::DisplayRole).toString();
        QAxObject* pCell = excel->getCellItem(pSheet,1,j+1);
        QAxObject* font = excel->getItemFont(pCell);
        font->setProperty("Bold", true);
        excel->setItemCenter(pCell);
        excel->setItemValue(pCell,value);
    }
    for(int i = 0;i<rows;++i){
        SetLoadingInfo(QString("正在处理,Excel 第%1行数据,共%2行").arg(i).arg(rows));
        for(int j = 0;j<cols;++j){
            if(mIsStop){
                return;
            }
            QString value = pModel->itemData(pModel->index(i,j)).value(Qt::DisplayRole).toString();//.->text();
            QVariant errType = pModel->itemData(pModel->index(i,j)).value(Qt::UserRole);
            if(!errType.isNull() && errType.toInt() == 1){
                //判断该值是否在正常范围内，超出范围则单元格颜色为红色
                QAxObject* pCell = excel->getCellItem(pSheet,i+2,j+1);
                excel->setItemBkColor(pCell,QColor(255,0,0,255));
                excel->setItemValue(pCell,value);
            }else{
                excel->setCell(pSheet,i+2,j+1,value);
            }
        }
    }
    excel->save();
}
