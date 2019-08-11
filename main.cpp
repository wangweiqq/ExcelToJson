/*电池壳孔数据Excel转换为Json文件*/
#include "dialog.h"
#include <QApplication>
#include <objbase.h>
int main(int argc, char *argv[])
{
   
    QApplication a(argc, argv);
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    Dialog w;
    w.show();
    return a.exec();
}
