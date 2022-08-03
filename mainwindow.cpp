#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "qDebug"
#include "QtWidgets"
#include <QMessageBox>


/*设置全局变量*/
QStandardItemModel * model = new QStandardItemModel;   //创建一个标准的条目模型
int speed_time = 10 ;//用于调整分流速度
QMap<QString, int> profession;//字典存储每个专业及对应专业设置的名额
Student_data *Stu[STUDENT_NUM]; //用于存储学生信息，及排序


/**************所做的合法性校验***********************
 * 1，对输入excel文件检测是否行数（除第一行）与学生人数相等
 * 2，对输入学生成绩校验是否为0<x<=100
 * 3，对输入专业校验是否为设定专业中的一个
 * 4，对设定校验是否学生人数等于专业数量总和
 * 5，对程序操作顺序校验是否先进行excel读取操作和输出目录选择操作
 * 6，校验是否读取了正确的excel文件
****************合法性校验***************************/



//构造函数
MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->tableView->setModel(model);
    setWindowTitle(TITLE);//设置程序标题

    this->ui->lineEdit->setText("欢迎使用专业分流系统-请读入excel文件");
    /*------------------------------------
     * ----将专业及对应人数存入字典------------
     * ----------------------------------*/
    profession.insert(SPECIALTY_1,SPECIALTY_NUM_1);
    profession.insert(SPECIALTY_2,SPECIALTY_NUM_2);
    profession.insert(SPECIALTY_3,SPECIALTY_NUM_3);
    profession.insert(SPECIALTY_4,SPECIALTY_NUM_4);
    profession.insert(SPECIALTY_5,SPECIALTY_NUM_5);
    profession.insert(SPECIALTY_6,SPECIALTY_NUM_6);
    profession.insert(SPECIALTY_7,SPECIALTY_NUM_7);

    /*------------------------------------
     * --------初始化UI的Label显示----------
     * ----------------------------------*/
    this->ui->label_1->setText(SPECIALTY_1);
    this->ui->label_2->setText(SPECIALTY_2);
    this->ui->label_3->setText(SPECIALTY_3);
    this->ui->label_4->setText(SPECIALTY_4);
    this->ui->label_5->setText(SPECIALTY_5);
    this->ui->label_6->setText(SPECIALTY_6);
    this->ui->label_7->setText(SPECIALTY_7);
    /*------------------------------------
     * --------初始化LCD专业名额显示---------
     * ----------------------------------*/
    this->ui->lcdNumber_1->display(SPECIALTY_NUM_1);
    this->ui->lcdNumber_2->display(SPECIALTY_NUM_2);
    this->ui->lcdNumber_3->display(SPECIALTY_NUM_3);
    this->ui->lcdNumber_4->display(SPECIALTY_NUM_4);
    this->ui->lcdNumber_5->display(SPECIALTY_NUM_5);
    this->ui->lcdNumber_6->display(SPECIALTY_NUM_6);
    this->ui->lcdNumber_7->display(SPECIALTY_NUM_7);
    this->ui->lcdNumber1->display(SPECIALTY_NUM_1);
    this->ui->lcdNumber2->display(SPECIALTY_NUM_2);
    this->ui->lcdNumber3->display(SPECIALTY_NUM_3);
    this->ui->lcdNumber4->display(SPECIALTY_NUM_4);
    this->ui->lcdNumber5->display(SPECIALTY_NUM_5);
    this->ui->lcdNumber6->display(SPECIALTY_NUM_6);
    this->ui->lcdNumber7->display(SPECIALTY_NUM_7);
    inform_input=0;
    inform_output=0;
    data_pre=0;
}


MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked(bool checked)
{
    if(checked==false){
         qDebug()<<"clicked";

         int check = SPECIALTY_NUM_1+SPECIALTY_NUM_2+SPECIALTY_NUM_3+SPECIALTY_NUM_4+SPECIALTY_NUM_5
                 +SPECIALTY_NUM_6+SPECIALTY_NUM_7;
         if(check != STUDENT_NUM)
         {
             if(inform_input==0){
                 QMessageBox::critical(this ,
                 "critical message" , "专业数量与学生人数不对等",
                 QMessageBox::Ok | QMessageBox::Default ,
                 QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
                 return;
             }
         }
         if(inform_input==0){
             QMessageBox::critical(this ,
             "critical message" , "请读入文件在进行后续操作",
             QMessageBox::Ok | QMessageBox::Default ,
             QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
             return;
         }
        if(inform_output==0){
             QMessageBox::critical(this ,
             "critical messag" , "请选择输出目录（空文件夹）",
             QMessageBox::Ok | QMessageBox::Default ,
             QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
             return;
        }
        if(data_pre==0){
             QMessageBox::critical(this ,
             "critical messag" , "文件读取有误请关闭后重新打开程序",
             QMessageBox::Ok | QMessageBox::Default ,
             QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
             return;
        }

         /*初始化专业队列，存储分配至每个专业的学生信息*/
         QQueue<Student_data*> class_1;
         QQueue<Student_data*> class_2;
         QQueue<Student_data*> class_3;
         QQueue<Student_data*> class_4;
         QQueue<Student_data*> class_5;
         QQueue<Student_data*> class_6;
         QQueue<Student_data*> class_7;
         /*初始化志愿队列，用于处理不同志愿*/
         QQueue<Student_data*> choices_2;
         QQueue<Student_data*> choices_3;
         QQueue<Student_data*> choices_4;
         QQueue<Student_data*> choices_5;
         QQueue<Student_data*> choices_6;
         QQueue<Student_data*> choices_7;
         QQueue<Student_data*> choices_adj;

         /********按照学分绩降序排序处理***********/
       // this->ui->lineEdit->setText("正在排序...");
      //   quickSort(Stu,0,STUDENT_NUM);

         /********开始分流***********/
        this->ui->lineEdit->setText("正在处理第一志愿...");
        initial_display();
        int count = 0;//在tableview的第几行显示数据

        /*************************************
         *************处理第一志愿**************
         ************* ***********************/
        for(int i=0; i< STUDENT_NUM; i++)
        {
            if(profession[Stu[i]->choice_1]>0)//该专业剩余人数大于0可以分配
            {
                ++count;
                --profession[Stu[i]->choice_1];
                Stu[i]->result = Stu[i]->choice_1;
                Stu[i]->result_num = "第一志愿录取";
                lcd_adj_display(Stu[i],&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
                display_data(Stu[i],count);
                QThread::msleep();
            }
            else{
            #if CHOICES_NUM>=2
                ++count;
                Stu[i]->result ="一志愿名额不足等待第二志愿处理";
                display_data(Stu[i],count);
                choices_2.enqueue(Stu[i]);//若设置学生志愿数大于2则存储choices_2中否则存入调剂队列中
                QThread::msleep();
            #else
                ++count;
                Stu[i]->result = "名额不足等待调剂";
                display_data(Stu[i],count);
                choices_adj.enqueue(Stu[i]);
                QThread::msleep();
            #endif
            }
        }

        Student_data *temp ;
        #if CHOICES_NUM>=2
        /*************************************
        *************处理第二志愿**************
        ************* ***********************/
        QThread::msleep(2000);
        model->clear();//清除viewtable显示的数据
        this->ui->lineEdit->setText("正在处理第二志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据


        while(choices_2.size()!=0)
        {
            temp = choices_2.dequeue();
            if(profession[temp->choice_2]>0){
                ++count;
                --profession[temp->choice_2];
                temp->result = temp->choice_2;
                temp->result_num = "第二志愿录取";
                lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
                display_data(temp,count);
                QThread::msleep();
            }
            else{
                #if CHOICES_NUM>=3
                    ++count;
                    temp->result = "二志愿名额不足等待三志愿处理";
                    display_data(temp,count);
                    choices_3.enqueue(temp);//若设置学生志愿数大于2则存储choices_2中否则存入调剂队列中
                    QThread::msleep();
                #else
                    ++count;
                    temp->result = "名额不足等待调剂";
                    display_data(temp,count);
                    choices_adj.enqueue(temp);
                    QThread::msleep();
                #endif
                }
        }
        QThread::msleep(2000);
        model->clear();//清除viewtable
        #endif

        #if CHOICES_NUM>=3
        /*************************************
        *************处理第三志愿**************
        ************* ***********************/
        this->ui->lineEdit->setText("正在处理第三志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据

        while(choices_3.size()!=0)
        {
             temp = choices_3.dequeue();
             if(profession[temp->choice_3]>0){
                 ++count;
                 --profession[temp->choice_3];
                 temp->result = temp->choice_3;
                 temp->result_num = "第三志愿录取";
                 lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
                 display_data(temp,count);
                 QThread::msleep();
             }
             else{
                #if CHOICES_NUM>=4
                    ++count;
                    temp->result = "三志愿名额不足等待四志愿处理";
                    display_data(temp,count);
                    choices_4.enqueue(temp);//若设置学生志愿数大于2则存储choices_2中否则存入调剂队列中
                    QThread::msleep();
                #else
                    ++count;
                    temp->result = "名额不足等待调剂";
                    display_data(temp,count);
                    choices_adj.enqueue(temp);
                    QThread::msleep();
                #endif
            }
        }
        #endif


        #if CHOICES_NUM>=4
        /*************************************
        *************处理第四志愿**************
        ************* ***********************/
        QThread::msleep(2000);
        model->clear();//清除viewtable
        this->ui->lineEdit->setText("正在处理第四志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据
        while(choices_4.size()!=0)
        {
             temp = choices_4.dequeue();
             if(profession[temp->choice_4]>0)
             {
                ++count;
                --profession[temp->choice_4];
                temp->result = temp->choice_4;
                temp->result_num = "第四志愿录取";
                lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
                display_data(temp,count);
                QThread::msleep();

             }else{
                #if CHOICES_NUM>=5
                    ++count;
                    temp->result = "四志愿名额不足等待五志愿处理";
                    display_data(temp,count);
                    choices_5.enqueue(temp);
                    QThread::msleep();
                #else
                    ++count;
                    temp->result = "名额不足等待调剂";
                    display_data(temp,count);
                    choices_adj.enqueue(temp);
                    QThread::msleep();
                #endif
                }
        }
        #endif

        #if CHOICES_NUM>=5
        /*************************************
        *************处理第五志愿**************
        ************* ***********************/
        QThread::msleep(2000);
        model->clear();//清除viewtable
        this->ui->lineEdit->setText("正在处理第五志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据
        while(choices_5.size()!=0)
        {
            temp = choices_5.dequeue();
            if(profession[temp->choice_5]>0)
            {
                ++count;
                --profession[temp->choice_5];
                temp->result = temp->choice_5;
                temp->result_num = "第五志愿录取";
                lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
                display_data(temp,count);
                QThread::msleep();
            }else{
            #if CHOICES_NUM>=6
                ++count;
                temp->result = "五志愿名额不足等待六志愿处理";
                display_data(temp,count);
                choices_6.enqueue(temp);
                QThread::msleep();
            #else
                ++count;
                temp->result = "名额不足等待调剂";
                display_data(temp,count);
                choices_adj.enqueue(temp);
                QThread::msleep();
            #endif
            }
        }
        #endif

        #if CHOICES_NUM>=6
        /*************************************
        *************处理第六志愿**************
        ************* ***********************/
        QThread::msleep(2000);
        model->clear();//清除viewtable
        this->ui->lineEdit->setText("正在处理第六志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据
        while(choices_6.size()!=0)
        {
            temp = choices_6.dequeue();
            if(profession[temp->choice_6]>0)
            {
                ++count;
                --profession[temp->choice_6];
                temp->result = temp->choice_6;
                temp->result_num = "第六志愿录取";
                lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
                display_data(temp,count);
                QThread::msleep();
            }else{
                #if CHOICES_NUM>=7
                    ++count;
                    temp->result = "六志愿名额不足等待七志愿处理";
                    display_data(temp,count);
                    choices_7.enqueue(temp);
                    QThread::msleep();
                #else
                    ++count;
                    temp->result = "名额不足等待调剂";
                    display_data(temp,count);
                    choices_adj.enqueue(temp);
                    QThread::msleep();
                #endif
            }
        }
        #endif

        #if CHOICES_NUM>=7
        /*************************************
        *************处理第七志愿**************
        ************* ***********************/
        QThread::msleep(2000);
        model->clear();//清除viewtable
        this->ui->lineEdit->setText("正在处理第七志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据
        while(choices_7.size()!=0)
        {
            temp = choices_7.dequeue();
            if(profession[temp->choice_7]>0)

        {
            ++count;
            --profession[temp->choice_7];
            temp->result = temp->choice_7;
            temp->result_num = "第七志愿录取";
            lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);
            display_data(temp,count);
            QThread::msleep();
        }else{
            ++count;
            temp->result = "名额不足等待调剂";
            display_data(temp,count);
            choices_adj.enqueue(temp);
            QThread::msleep();
            }
        }
        #endif

        /*************************************
        *************处理调剂志愿***************
        ************* ***********************/
        QThread::msleep(2000);
        model->clear();//清除viewtable
        this->ui->lineEdit->setText("正在处理调剂志愿...");
        initial_display();
        count = 0;//在tableview的第几行显示数据,初始化为-1
        QString sorted[7]={SPECIALTY_1,SPECIALTY_2,SPECIALTY_3,SPECIALTY_4,SPECIALTY_5
                       ,SPECIALTY_6,SPECIALTY_7}; //基于每个专业剩余数量，从大到小存放专业
        QString tep_sort;
        for(int c = 0 ;c < 7 ;c++){
            for(int h = c+1 ;h < 7 ;h++){
                if(profession[sorted[c]] < profession[sorted[h]]){
                    tep_sort = sorted[c];
                    sorted[c] = sorted[h];
                    sorted[h] = tep_sort;
                }
            }
        }
        int state = 0;
        while(choices_adj.size() != 0){
            temp = choices_adj.dequeue();//取除队列顶学生数据
            qDebug() << "调剂";
            if(profession[sorted[0]] == 0 && profession[sorted[1]] == 0
                    && profession[sorted[2]] == 0 && profession[sorted[3]] == 0
                    && profession[sorted[4]] == 0 && profession[sorted[5]] == 0
                    && profession[sorted[6]] == 0){
                qDebug() << "调剂error";//学生总人数大于专业设置总数量时break；
                break;
            }
            while(1){
                if(state == 0)
                {
                    if(profession[sorted[0]] != 0){
                        state = 1;//状态转移
                        --profession[sorted[0]];//该专业名额减1
                        ++count;//viewtable显示行数加1
                        temp->result = sorted[0];//将专业分配结果存入学生结构体
                        temp->result_num = "调剂录取"; //存储该生来自调剂志愿
                        break;
                    }
                    else{
                        state = 1;
                    }
                }
                else if(state == 1)
                {
                    if(profession[sorted[1]] != 0)
                    {
                        state = 2;
                        --profession[sorted[1]];
                        ++count;
                        temp->result = sorted[1];
                        temp->result_num = "调剂录取";
                        break;
                    }
                    else{
                        state = 2;
                    }
                }
                else if(state == 2)
                {
                    if(profession[sorted[2]] != 0)
                    {
                        state = 3;
                        --profession[sorted[2]];
                        ++count;
                        temp->result = sorted[2];
                        temp->result_num = "调剂录取";
                        break;
                    }
                    else
                        state = 3;
                }
                else if(state == 3)
                {
                    if(profession[sorted[3]])
                    {
                        state = 4;
                        --profession[sorted[3]];
                        ++count;
                        temp->result = sorted[3];
                        temp->result_num = "调剂录取";
                        break;
                    }
                    else
                        state = 4;
                }
                else if(state == 4)
                {
                    if(profession[sorted[4]] != 0){
                        state = 5;
                        --profession[sorted[4]];
                        ++count;
                        temp->result = sorted[4];
                        temp->result_num = "调剂录取";
                        break;
                    }
                    else
                        state = 5;
                }
                else if(state == 5)
                {
                    if(profession[sorted[5]] != 0)
                    {
                        state = 6;
                        --profession[sorted[5]];
                        ++count;
                        temp->result = sorted[5];
                        temp->result_num = "调剂录取";
                        break;
                    }
                    else
                        state = 6;
                }
                else if(state == 6)
                {
                    if(profession[sorted[6]] != 0)
                    {
                        state = 0;
                        --profession[sorted[6]];
                        ++count;
                        temp->result = sorted[6];
                        temp->result_num = "调剂录取";
                        break;
                    }
                    else
                        state = 0;
                }
            }
            lcd_adj_display(temp,&class_1,&class_2,&class_3, &class_4, &class_5, &class_6, &class_7);//lcd显示
            display_data(temp,count);//tableview显示
            QThread::msleep();
            if(choices_adj.size()==1)
                this->ui->lineEdit->setText("分流完成正在保存结果...");
        }
#if SAVE_RESULT==1
        save_result(&class_1,&class_2,&class_3,&class_4,&class_5,&class_6,&class_7);
#endif
        this->ui->lineEdit->setText("结果已保存到指定目录下，谢谢使用");

    }

}
















void MainWindow::on_action_excel_triggered()
{
    QString input_file;
    input_file = QFileDialog::getOpenFileName(this,tr("打开excel文件"),"C:\Users",tr("(*.xlsx)"));

        if(QString(input_file).isEmpty()==false){
            if(inform_input==1){
                QMessageBox::critical(this ,
                "critical message" , "请勿重复读取文件",
                QMessageBox::Ok | QMessageBox::Default ,
                QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
                return;
            }
            qDebug()<<input_file;
            readEnvXlsFile(input_file);
            inform_input=1;
        }
        if(data_pre == 1)
            this->ui->lineEdit->setText("学生信息读取完毕-请选择输出目录（空文件夹）");
}



//输出目录
void MainWindow::on_action_triggered()
{


    output_dic = QFileDialog::getExistingDirectory();
    if(QString(output_dic).isEmpty()==false)
    {
         qDebug()<<output_dic;
         qDebug()<< profession.values();
         inform_output = 1;
         this->ui->lineEdit->setText("输出目录设置完成-请点击start进行分流");
    }

}


//处理速度设置
void MainWindow::on_horizontalSlider_sliderMoved(int position)
{
    qDebug()<<position;
     = position*10;
}


//读取excel
void MainWindow::readEnvXlsFile(QString FileName)
{

    QAxObject *excel = NULL;
    QAxObject *workbooks = NULL;
    QAxObject *workbook = NULL;

    excel = new QAxObject("Excel.Application");
    if (!excel)
    {
         qDebug()<<"EXCEL对象丢失!";
    }
    excel->dynamicCall("SetVisible(bool)", false);
    workbooks = excel->querySubObject("WorkBooks");
    workbook = workbooks->querySubObject("Open(QString, QVariant)", FileName);
    QAxObject * worksheet = workbook->querySubObject("WorkSheets(int)", 1);//打开第一个sheet
    QAxObject * usedrange = worksheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
    QAxObject * rows = usedrange->querySubObject("Rows");
    QAxObject * columns = usedrange->querySubObject("Columns");
    int intRows = rows->property("Count").toInt();
    int intCols = columns->property("Count").toInt();
    int c = intRows-1;
    intCols;
    qDebug() << intCols;
    if(c!=STUDENT_NUM)
    {
        QMessageBox::critical(this ,
        "critical message" , "文件除第一行行数与学生数需对应",
        QMessageBox::Ok | QMessageBox::Default ,
        QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
        return;
    }

    //批量载入数据，这里默认读取A2:K从第二行到最后，第一列到最后A学号B姓名CD成绩E-K志愿
    QString Range = "A2:K" +QString::number(intRows);
    QAxObject *allEnvData = worksheet->querySubObject("Range(QString)", Range);
    QVariant allEnvDataQVariant = allEnvData->property("Value");
    QVariantList allEnvDataList = allEnvDataQVariant.toList();
    QString bug;
    for(int i=0; i < intRows-1; i++)
    {
        Stu[i] = new Student_data;

        QVariantList allEnvDataList_i =  allEnvDataList[i].toList() ;

        //读入学号
        Stu[i]->sno = allEnvDataList_i[0].toString();

        //读入姓名
        Stu[i]->name = allEnvDataList_i[1].toString();

        //读入学分绩1------这里改成读取排名了
        if(0<allEnvDataList_i[2].toFloat()&&allEnvDataList_i[2].toFloat()<=1000){
            Stu[i]->credits_1 = allEnvDataList_i[2].toFloat();
            qDebug()<< Stu[i]->credits_1;
        }
        else{
            bug = Stu[i]->sno + "的排名输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }

        #if GRADE_NUM>1
        //读入学分绩2
            if(0<allEnvDataList_i[3].toFloat()&&allEnvDataList_i[3].toFloat()<=100){
                Stu[i]->credits_2 = allEnvDataList_i[3].toFloat();
            }
            else{
                bug = Stu[i]->sno + "的学分绩2输入有误";
                QMessageBox::critical(this ,
                "critical message" , bug,
                QMessageBox::Ok | QMessageBox::Default ,
                QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
                return;
            }
        #endif

        /**********读入第一志愿***********/
        if(profession.count(allEnvDataList_i[4].toString())!=0){
            Stu[i]->choice_1 = allEnvDataList_i[4].toString();
        }
        else{
            bug = Stu[i]->sno + "的第1志愿输入有误";
                QMessageBox::critical(this ,
                "critical message" , bug,
                QMessageBox::Ok | QMessageBox::Default ,
                QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
                return;
        }

        /**********读入第二志愿***********/
        #if CHOICES_NUM>=2
        if(profession.count(allEnvDataList_i[5].toString())!=0){
            Stu[i]->choice_2 = allEnvDataList_i[5].toString();
            }
        else{
            bug = Stu[i]->sno + "的第2志愿输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }
        #endif

        /**********读入第三志愿***********/
        #if CHOICES_NUM>=3
        if(profession.count(allEnvDataList_i[6].toString())!=0){
            qDebug()<<allEnvDataList_i[6].toString();
            Stu[i]->choice_3 = allEnvDataList_i[6].toString();
            }
        else{
            bug = Stu[i]->sno + "的第3志愿输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }
        #endif

        /**********读入第四志愿***********/
        #if CHOICES_NUM>=4
        if(profession.count(allEnvDataList_i[7].toString())!=0){
            Stu[i]->choice_4 = allEnvDataList_i[7].toString();
            }
        else{
            bug = Stu[i]->sno + "的第4志愿输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }
        #endif

        #if CHOICES_NUM>=5
        /**********读入第五志愿***********/
        if(profession.count(allEnvDataList_i[8].toString())!=0){
            qDebug()<<allEnvDataList_i[8].toString();
            Stu[i]->choice_5 = allEnvDataList_i[8].toString();
            }
        else{
            bug = Stu[i]->sno + "的第5志愿输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }
        #endif

        #if CHOICES_NUM>=6
        /**********读入第六志愿***********/
        if(profession.count(allEnvDataList_i[9].toString())!=0){
            qDebug()<<allEnvDataList_i[9].toString();
            Stu[i]->choice_6 = allEnvDataList_i[9].toString();
            }
        else{
            bug = Stu[i]->sno + "的第6志愿输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }
        #endif

        #if CHOICES_NUM>=7
        /**********读入第七志愿***********/
        if(profession.count(allEnvDataList_i[10].toString())!=0){
            qDebug()<<allEnvDataList_i[10].toString();
            Stu[i]->choice_7 = allEnvDataList_i[10].toString();
            }
        else{
            bug = Stu[i]->sno + "的第7志愿输入有误";
            QMessageBox::critical(this ,
            "critical message" , bug,
            QMessageBox::Ok | QMessageBox::Default ,
            QMessageBox::Cancel | QMessageBox::Escape , 	0 );//错误信息提示
            return;
        }
        #endif

        }
    data_pre = 1;
    workbooks->dynamicCall("Close()");
}



#if GRADE_NUM==1

/**************一个学分绩的排序算法**************/
int MainWindow::quickSortPartition(Student_data *s[], int l, int r) {
    int i = l, j = r;
    Student_data *x = s[l];
        while (i < j)
        {
            while (i < j && s[j]->credits_1 >= x->credits_1)
                j--;
            if (i < j)
                s[i++] = s[j];
            while (i < j && s[i]->credits_1 <= x->credits_1)
                i++;
            if (i < j)
                s[j--] = s[i];
        }
        s[i] = x;
        return i;
}
void MainWindow::quickSort(Student_data *s[], int l, int r)
{

    if (l >= r) {
        return;
    }
    int i = quickSortPartition(s, l, r);
    quickSort(s, l, i - 1);
    quickSort(s, i + 1, r);
}


#else
/**************两个学分绩的排序算法**************/
void MainWindow::quickSort(Student_data *s[], int l, int r)
{   l;
    Student_data *temp;
    for (int i = 0; i <= r; i++) {
        for (int j = i+1; j <= r; j++) {
            if (s[i]->credits_1 <= s[j]->credits_1)
            {
                if (s[i]->credits_1 != s[j]->credits_1)
                {
                    temp = s[i];
                    s[i] = s[j];
                    s[j] = temp;
                }
                else
                {
                    if (s[i]->credits_2 <= s[j]->credits_2) {
                        temp = s[i];
                        s[i] = s[j];
                        s[j] = temp;
                    }
                    else
                        continue;
                }
            }
        }
    }
}
#endif

void MainWindow::initial_display(){
    model->setHorizontalHeaderItem(0, new QStandardItem("学号") );
    this->ui->tableView->setColumnWidth(0, 300);    //设置列的宽度

    model->setHorizontalHeaderItem(1, new QStandardItem("姓名"));
    this->ui->tableView->setColumnWidth(1, 300);

    model->setHorizontalHeaderItem(2, new QStandardItem("所分配专业"));
    this->ui->tableView->setColumnWidth(2, 300);
}


void MainWindow::display_data(Student_data *s,int count)
{
    count = count - 1;
     model->setItem(count, 0, new QStandardItem(s->sno) );
     model->item(count,0)->setTextAlignment(Qt::AlignCenter); //设置居中

     model->setItem(count, 1, new QStandardItem(s->name) );
     model->item(count, 1)->setTextAlignment(Qt::AlignCenter);

     model->setItem(count,2, new QStandardItem(s->result) );
     model->item(count, 2)->setTextAlignment(Qt::AlignCenter);

     model->item(count, 1)->setFont( QFont( "Times", 10, QFont::Black ) ); //设置字体
     model->item(count, 0)->setForeground(QBrush(QColor(255, 0, 0)));//设置单元格文本颜色

    //滚动条到最后一行
    this->ui->tableView->scrollToBottom();
    //动态显示信息
    qApp->processEvents();

 }


void MainWindow::lcd_adj_display(Student_data *temp
                                 , QQueue<Student_data*> *class_1
                                 , QQueue<Student_data*> *class_2
                                 , QQueue<Student_data*> *class_3
                                 , QQueue<Student_data*> *class_4
                                 , QQueue<Student_data*> *class_5
                                 , QQueue<Student_data*> *class_6
                                 , QQueue<Student_data*> *class_7)
{
    if(temp->result==SPECIALTY_1){
        class_1->enqueue(temp);
        this->ui->lcdNumber1->display(profession[temp->result]);
    }
    else if(temp->result==SPECIALTY_2){
        class_2->enqueue(temp);
        this->ui->lcdNumber2->display(profession[temp->result]);
    }
    else if(temp->result==SPECIALTY_3){
        class_3->enqueue(temp);
        this->ui->lcdNumber3->display(profession[temp->result]);
    }
    else if(temp->result==SPECIALTY_4){
        class_4->enqueue(temp);
        this->ui->lcdNumber4->display(profession[temp->result]);
    }
    else if(temp->result==SPECIALTY_5){
        class_5->enqueue(temp);
        this->ui->lcdNumber5->display(profession[temp->result]);
    }
    else if(temp->result==SPECIALTY_6){
        class_6->enqueue(temp);
        this->ui->lcdNumber6->display(profession[temp->result]);
    }
    else if(temp->result==SPECIALTY_7){
        class_7->enqueue(temp);
        this->ui->lcdNumber7->display(profession[temp->result]);
    }
    else
        qDebug() << "error";

}



/*****************数据分析与结果保存**********************/
void MainWindow::save_result(QQueue<Student_data*> *class_1,
                            QQueue<Student_data*> *class_2,
                            QQueue<Student_data*> *class_3,
                            QQueue<Student_data*> *class_4,
                            QQueue<Student_data*> *class_5,
                            QQueue<Student_data*> *class_6,
                            QQueue<Student_data*> *class_7)
{   /*****************统计信息******************/
    QMap<QString, int> data_ana;
    data_ana.insert("第一志愿录取",0);
    data_ana.insert("第二志愿录取",0);
    data_ana.insert("第三志愿录取",0);
    data_ana.insert("第四志愿录取",0);
    data_ana.insert("第五志愿录取",0);
    data_ana.insert("第六志愿录取",0);
    data_ana.insert("第七志愿录取",0);
    data_ana.insert("调剂录取",0);
    float total_grade1 = 0;
    float total_grade2 = 0;
    /*****************保存总体结果******************/
    int count_r= 1;
    QString fileName = output_dic+"/"+"分流结果"+".xlsx";
    newExcel(fileName);
    setCellValue(1,count_r++,"学号");
    setCellValue(1,count_r++,"姓名");
    setCellValue(1,count_r++,"排名");
    #if GRADE_NUM==2
    setCellValue(1,count_r++,"成绩2");
    #endif
    setCellValue(1,count_r++,"分配专业");
    setCellValue(1,count_r,"第几志愿录取");
    for(int i=0;i<STUDENT_NUM;i++)
    {
        count_r=1;
        setCellValue(i+2,count_r++,Stu[i]->sno);
        setCellValue(i+2,count_r++,Stu[i]->name);
        setCellValue(i+2,count_r++,QString("%1").arg(Stu[i]->credits_1));
        total_grade1 = Stu[i]->credits_1 + total_grade1 ; //统计成绩信息
        #if GRADE_NUM==2
        setCellValue(i+2,count_r++,QString("%1").arg(Stu[i]->credits_2));
        total_grade2 = Stu[i]->credits_2 + total_grade2 ;  //统计成绩信息
        #endif
        setCellValue(i+2,count_r++,Stu[i]->result);
        setCellValue(i+2,count_r,Stu[i]->result_num);
        data_ana[Stu[i]->result_num]++;
    }
    int count_t=STUDENT_NUM+3;
    /**成绩1**/
    total_grade1 = total_grade1/STUDENT_NUM;
    setCellValue(count_t,1,"平均排名");
    setCellValue(count_t++,2,QString("%1").arg(total_grade1));
    /**成绩2**/
    #if GRADE_NUM==2
    total_grade2 = total_grade2/STUDENT_NUM;
    setCellValue(count_t,1,"平均成绩2");
    setCellValue(count_t++,2,QString("%1").arg(total_grade2));
    #endif
    setCellValue(count_t,1,"第一志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第一志愿录取"]));
    #if CHOICES_NUM>=2
    setCellValue(count_t,1,"第二志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第二志愿录取"]));
    #endif
    #if CHOICES_NUM>=3
    setCellValue(count_t,1,"第三志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第三志愿录取"]));
    #endif
    #if CHOICES_NUM>=4
    setCellValue(count_t,1,"第四志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第四志愿录取"]));
    #endif
    #if CHOICES_NUM>=5
    setCellValue(count_t,1,"第五志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第五志愿录取"]));
    #endif
    #if CHOICES_NUM>=6
    setCellValue(count_t,1,"第六志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第六志愿录取"]));
    #endif
    #if CHOICES_NUM>=7
    setCellValue(count_t,1,"第七志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第七志愿录取"]));
    #endif

    setCellValue(count_t,1,"调剂人数");
    setCellValue(count_t,2,QString("%1").arg(data_ana["调剂录取"]));

    saveExcel(fileName);
    freeExcel();



    /*****************保存专业1结果******************/
    save_profession_result(SPECIALTY_1,class_1);
    /*****************保存专业2结果******************/
    save_profession_result(SPECIALTY_2,class_2);
    /*****************保存专业3结果******************/
    save_profession_result(SPECIALTY_3,class_3);
    /*****************保存专业4结果******************/
    save_profession_result(SPECIALTY_4,class_4);
    /*****************保存专业5结果******************/
    save_profession_result(SPECIALTY_5,class_5);
    /*****************保存专业6结果******************/
    save_profession_result(SPECIALTY_6,class_6);
    /*****************保存专业7结果******************/
    save_profession_result(SPECIALTY_7,class_7);
}



// 新建一个excel
void MainWindow::newExcel(const QString &fileName)
{
    pApplication = new QAxObject("Excel.Application");
    if (pApplication == nullptr) {
        qWarning("pApplication\n");
        return;
    }
    pApplication->dynamicCall("SetVisible(bool)",false);// false不显示窗体
    pApplication->setProperty("DisplayAlerts",false);// 不显示任何警告信息
    pWorkBooks = pApplication->querySubObject("Workbooks");
    QFile file(fileName);
    if (file.exists()) {
        pWorkBook = pWorkBooks->querySubObject("Open(const QString&)",fileName);
    } else {
        pWorkBooks->dynamicCall("Add");
        pWorkBook = pApplication->querySubObject("ActiveWorkBook");
    }
    pSheets = pWorkBook->querySubObject("Sheets");
    pSheet = pSheets->querySubObject("Item(int)",1);
}


// 增加一个worksheet
void MainWindow::appendSheet(const QString &sheetName,int cnt)
{
    QAxObject *pLastSheet = pSheets->querySubObject("Item(int)",cnt);
    pSheets->querySubObject("Add(QVariant)",pLastSheet->asVariant());
    pSheet = pSheets->querySubObject("Item(int)",cnt);
    pLastSheet->dynamicCall("Move(QVariant)",pSheet->asVariant());
    pSheet->setProperty("Name",sheetName);
}

// 向Excel单元格中写入数据
void MainWindow::setCellValue(int row,int column,const QString &value)
{
    QAxObject *pRange = pSheet->querySubObject("Cells(int,int)",row,column);
    pRange->dynamicCall("Value",value);
}

// 保存excel
void MainWindow::saveExcel(const QString &fileName)
{
    pWorkBook->dynamicCall("SaveAs(const QString &)",QDir::toNativeSeparators(fileName));
}

// 释放excel
void MainWindow::freeExcel()
{
    if (pApplication != nullptr) {
        pApplication->dynamicCall("Quit()");
        delete pApplication;
        pApplication = nullptr;
    }
}

//保存专业结果
void MainWindow::save_profession_result(QString spec,QQueue<Student_data*> *class_){
    /***********统计信息*************/
    QMap<QString, int> data_ana;
    data_ana.insert("第一志愿录取",0);
    data_ana.insert("第二志愿录取",0);
    data_ana.insert("第三志愿录取",0);
    data_ana.insert("第四志愿录取",0);
    data_ana.insert("第五志愿录取",0);
    data_ana.insert("第六志愿录取",0);
    data_ana.insert("第七志愿录取",0);
    data_ana.insert("调剂录取",0);
    float total_grade1 = 0;
    float total_grade2 = 0;

    Student_data *temp ;
    int count_r=1;
    int i = 0;
    QString fileName1 = output_dic+"/"+spec+".xlsx";
    newExcel(fileName1);
    setCellValue(1,count_r++,"学号");
    setCellValue(1,count_r++,"姓名");
    setCellValue(1,count_r++,"排名");
    #if GRADE_NUM==2
    setCellValue(1,count_r++,"成绩2");
    #endif
    setCellValue(1,count_r,"第几志愿录取");
    while(class_->size()!=0)
    {
        temp = class_->dequeue();
        count_r=1;
        setCellValue(i+2,count_r++,temp->sno);
        setCellValue(i+2,count_r++,temp->name);
        setCellValue(i+2,count_r++,QString("%1").arg(temp->credits_1));
        total_grade1 = temp->credits_1 + total_grade1 ; //统计成绩信息
        #if GRADE_NUM==2
        setCellValue(i+2,count_r++,QString("%1").arg(temp->credits_2));
        total_grade2 = temp->credits_2 + total_grade2 ; //统计成绩信息
        #endif
        setCellValue(i+2,count_r,temp->result_num);
        data_ana[temp->result_num]++;
        i++;

    }
    int count_t=i+3;
    /**成绩1**/
    total_grade1 = total_grade1/i;
    setCellValue(count_t,1,"平均排名");
    setCellValue(count_t++,2,QString("%1").arg(total_grade1));
    /**成绩2**/
    #if GRADE_NUM==2
    total_grade2 = total_grade2/i;
    setCellValue(count_t,1,"平均成绩2");
    setCellValue(count_t++,2,QString("%1").arg(total_grade2));
    #endif
    setCellValue(count_t,1,"第一志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第一志愿录取"]));
    #if CHOICES_NUM>=2
    setCellValue(count_t,1,"第二志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第二志愿录取"]));
    #endif
    #if CHOICES_NUM>=3
    setCellValue(count_t,1,"第三志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第三志愿录取"]));
    #endif
    #if CHOICES_NUM>=4
    setCellValue(count_t,1,"第四志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第四志愿录取"]));
    #endif
    #if CHOICES_NUM>=5
    setCellValue(count_t,1,"第五志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第五志愿录取"]));
    #endif
    #if CHOICES_NUM>=6
    setCellValue(count_t,1,"第六志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第六志愿录取"]));
    #endif
    #if CHOICES_NUM>=7
    setCellValue(count_t,1,"第七志愿人数");
    setCellValue(count_t++,2,QString("%1").arg(data_ana["第七志愿录取"]));
    #endif
    setCellValue(count_t,1,"调剂人数");
    setCellValue(count_t,2,QString("%1").arg(data_ana["调剂录取"]));

    saveExcel(fileName1);
    freeExcel();
}
