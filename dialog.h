#ifndef DIALOG_H
#define DIALOG_H

#include <QDialog>
#include <QVector>
#include <string>
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
    int GetFileContent(QString filename);
public slots:
    void on_btnOpenExcel_clicked();
private:
    Ui::Dialog *ui;
};

#endif // DIALOG_H
