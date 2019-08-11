#ifndef DIALOG_H
#define DIALOG_H

#include <QDialog>
#include <QVector>
#include <string>
#include <QStandardItemModel>
#include <QItemSelectionModel>
using namespace std;
namespace Ui {
class Dialog;
}

class Dialog : public QDialog
{
    Q_OBJECT

public:
    explicit Dialog(QWidget *parent = 0);
    ~Dialog();
    //bool xlsReader(QString excelPath,vector<string> &stuNames);
    void GetFileContent(QString filename);
public slots:
    void on_btnOpenExcel_clicked();
private:
    Ui::Dialog *ui;
    QStandardItemModel* pModel;
    QItemSelectionModel* pSelection;
};

#endif // DIALOG_H
