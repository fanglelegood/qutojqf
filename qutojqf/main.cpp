#include "mainwindow.h"

#include <QApplication>


int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();

    return a.exec();

//    //初始化python模块
//           Py_Initialize();
//           if ( !Py_IsInitialized() )
//           {
//            return -1;
//           }
//           PyRun_SimpleString("import sys");
//           PyRun_SimpleString("sys.argv = ['python.py']");
//           PyRun_SimpleString("sys.path.append('./')");

//           //导入scriptSecond.py模块
//           PyObject* pModule = PyImport_ImportModule("scriptSecond");
//           if (!pModule) {
//                   printf("Cant open python file!\n");
//                   return -1;
//               }
//           //获取scriptSecond模块中的temperImg函数
//          PyObject* pFunhello= PyObject_GetAttrString(pModule,"temperImg");
//          //注释掉的这部分是另一种获得scriptSecond模块中的temperImg函数的方法
//      //    PyObject* pDict = PyModule_GetDict(pModule);
//      //       if (!pDict) {
//      //           printf("Cant find dictionary.\n");
//      //           return -1;
//      //       }
//      //    PyObject* pFunhello = PyDict_GetItemString(pDict, "temperImg");

//           if(!pFunhello){
//               cout <<"Get function hello failed"<< endl;
//               return -1;
//           }
//           //调用temperImg函数
//           PyObject_CallFunction(pFunhello,NULL);
//           //结束,释放python
//           Py_Finalize();



//       qDebug() << pReturn << Qt::endl;

//            int SizeOfList = PyList_Size(pReturn);//List对象的大小，这里SizeOfList =
//            for(int i = 0; i < SizeOfList; i++){
//                PyObject *Item = PyList_GetItem(pReturn, i);//获取List对象中的每一个元素
//                int result;
//                PyArg_Parse(Item, "c", &result);//s表示转换成string型变量
//              QString result = PyBytes_AsString(Item);
//                qDebug() << result << Qt::endl; //输出元素
//                Py_DECREF(Item); //释放空间
//             }





}
