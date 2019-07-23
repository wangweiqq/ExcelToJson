#ifndef CEXPORTEXCEL_H
#define CEXPORTEXCEL_H

#include <QObject>
#include <QTableView>
#include <QThread>
#include "exceloperator.h"
class CExportExcel : public QThread
{
    Q_OBJECT
public:
    explicit CExportExcel(const QAbstractItemView& v,const QString filePath,const QString sheetName);
    ~CExportExcel();
protected:
    void run() override;
public:
    /*设置保存名*/
    void SetFileName(QString file);
signals:
    void SetLoadingInfo(QString);
public slots:


private:
    const QAbstractItemView& view;
    ExcelOperator* excel;
    /*保存文件名*/
    const QString mFilePath;
    /*Sheet表名称*/
    const QString mSheetName;
    /*是否要停止线程*/
    bool mIsStop = false;
};

#endif // CEXPORTEXCEL_H
