# <b style="color:blue;">使用说明</b>

1. 把从学生收集到的包含**通用答题卡.xlsx**文件的文件夹放在本目录下的folder文件夹中,程序会自动搜索**older**下所有.xlsx文件,不要在<kbd>folder</kbd>文件夹中放非**通用答题卡.xls**模板的其它.xlsx文件。

2. 确保系统中已经正确安装了Python3。

3. 运行**收集信息.bat** 。 

4. 收集到的答题卡信息将会存放在**out**文件夹**统计.xlsx**中，可多次执行**收集信息.bat** 记录会追加到**统计.xlsx**中。

5. 如果要重新生成新的文件，请删除**统计.xlsx**后再执行**收集信息.bat**。

6. 执行**收集信息.bat**之前，确保将要被读取或写入的文件是关闭状态，否则读取或写入程序会失败。
7. 每一次执行的日志存放在**out/log.txt**文件中。
8. Folder文件夹中有**通用答题卡.xlsx**的样板。