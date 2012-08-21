Autolecture, for BJFU
=====================
根据BJFU课表信息，通过win32com生成outlook appointment，以同步至S60v3手机，作课程提醒用。

TODO
----
目前项目以基本完成，debug语句仍保留在源文件中，生成outlook appointment的call被注释。
*   修改appointment的recurrence类型为每周发生，使一门课程成为一个系列约会
*   修改公用常量的提供方式
*   完善docstring及注释
*   完善代码风格