#ifndef CONFIG_H
#define CONFIG_H
#pragma execution_character_set("utf-8")
/*excel说明
 * 列说明：
 * 1，excel共计11列
 * 2，excel第一列为学号
 * 3，excel第二列为姓名
 * 4，excel第三列为第一学分绩第四列为第二学分绩（第三列成绩优先级高于第四列）。如，每个同学只有一个学分绩，则将其填入第三列，第四列不用填任何数据，直接从第五列开始填写学生志愿。
 * 如果有两个学分绩时先比较优先级高的学分绩（第三列），若相同则比较优先级低（第四列）的学分绩
 * 5，excel第五到十二列为学生志愿1-7（由左到右，优先级高->低）。如，每个同学只有3志愿，则按照优先级从第五列到第七列填入对应志愿，八到十二列不用填任何数据。
 * 行说明
 * 6，excel第一行为对该列的信息说明
 * 7，除第一行外，行数等于学生人数
务必严格按照此格式制作excel文件*/
/************************************以下为需要设置的地方******************************************/
/*设置说明
 * 1，设置的分流学生人数应与各专业设置人数之和相等
 * 2，输入文件确保学生以进行升序排名
 * 3，学生专业志愿可选择1-7个，本次学生志愿为3个。
 * 4，设置内容应该与excel对应
*/
#define TITLE "学生分流系统 1.0--作者：曹峻凡 指导老师：张向利"  //程序的名称，可根据需要修改

#define STUDENT_NUM 716     //设置分流的学生人数
#define GRADE_NUM 1     //学分绩数量，可选1或2---PS不用设置这个，这一版本程序内部没有对学生排序
#define CHOICES_NUM 3   //学生专业志愿数量，可选1-7

#define SAVE_RESULT 1   //1为将结果保存，0为不保存

/**********设置专业1名称及对应名额**********/
#define SPECIALTY_1 "6.通信工程卓越班"  //名称改双引号内
#define SPECIALTY_NUM_1 29     //名额

/**********设置专业2名称及对应名额**********/
#define SPECIALTY_2 "7.电子信息工程卓越班"
#define SPECIALTY_NUM_2 29

/**********设置专业3名称及对应名额**********/
#define SPECIALTY_3 "1.通信工程"
#define SPECIALTY_NUM_3 172

/**********设置专业4名称及对应名额**********/
#define SPECIALTY_4 "2.电子信息工程"
#define SPECIALTY_NUM_4 172

/**********设置专业5名称及对应名额**********/
#define SPECIALTY_5 "3.电子科学与技术"
#define SPECIALTY_NUM_5 143

/**********设置专业6名称及对应名额**********/
#define SPECIALTY_6 "4.微电子科学与技术"
#define SPECIALTY_NUM_6 108

/**********设置专业7名称及对应名额**********/
#define SPECIALTY_7 "5.导航工程"
#define SPECIALTY_NUM_7 64

/************************************以上为需要设置的地方******************************************/

/*存储学生信息*/
typedef struct Student
{
    QString sno;    //第一列，学号
    QString name;   //第二列，姓名
    float credits_1;    //第三列，学分绩1
    float credits_2;    //第四列，学分绩2
    QString choice_1;   //第五列，第一志愿
    QString choice_2;   //第六列，第二志愿
    QString choice_3;
    QString choice_4;
    QString choice_5;
    QString choice_6;
    QString choice_7;   //第十一列，第七志愿
    QString result; //为该生分配的专业
    QString result_num;//分配的是第几志愿
} Student_data;


#endif // CONFIG_H
