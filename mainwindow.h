#ifndef MAINWINDOW_H
#define MAINWINDOW_H


#include <QMainWindow>
#include "procession.h"
#include<ActiveQt/QAxObject>


namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

public:
    void initial_display(); //tableview初始化

    void display_data(Student_data *s,int count); //tableview显示一行数据

    void lcd_adj_display(Student_data *temp,QQueue<Student_data*> *class_1,QQueue<Student_data*> *class_2,
                    QQueue<Student_data*> *class_3,QQueue<Student_data*> *class_4,QQueue<Student_data*> *class_5
                    ,QQueue<Student_data*> *class_6,QQueue<Student_data*> *class_7);//调剂时LCD显示并将学生信息存入对应专业结构体

    void readEnvXlsFile(QString FileName);//将文件读取到结构体中

    void quickSort(Student_data *s[], int l, int r);//排序

    int quickSortPartition(Student_data *s[], int l, int r);//排序

    /**********************************生成excel相关处理********************************/
    void save_result(QQueue<Student_data*> *class_1,
                                QQueue<Student_data*> *class_2,
                                QQueue<Student_data*> *class_3,
                                QQueue<Student_data*> *class_4,
                                QQueue<Student_data*> *class_5,
                                QQueue<Student_data*> *class_6,
                                QQueue<Student_data*> *class_7);

    void newExcel(const QString &fileName);

    void appendSheet(const QString &sheetName,int cnt);

    void setCellValue(int row,int column,const QString &value);

    void saveExcel(const QString &fileName);

    void freeExcel();

    void save_profession_result(QString spec,QQueue<Student_data*> *class_);

private:
    QAxObject *pApplication;
    QAxObject *pWorkBooks;
    QAxObject *pWorkBook;
    QAxObject *pSheets;
    QAxObject *pSheet;
    QString output_dic;
    int inform_input;
    int inform_output;
    int data_pre;

private slots:
    void on_pushButton_clicked(bool checked);

    void on_action_excel_triggered();

    void on_action_triggered();

    void on_horizontalSlider_sliderMoved(int position);

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H



