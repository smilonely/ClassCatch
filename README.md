# QCU 课程表抓取

*目前仅更新了 xls 课程表抓取程序，抓取网页的版本何时上传还是个未知数 ：）*

**项目思路：抓取课程表之后生成日历文件`.ics`，之后可以导入设备的系统日历，抛弃全是广告的「超级课程表」之流。**

   

## 关于 `.ics` 文件

`ics` 是 iCalendar 文件的后缀名，iCalendar 是「日历数据交换」的标准。

总而言之，有了它你可以很方便地将事件和提醒导入系统日历应用。

   

### 什么是 iCalendar

- 参阅 [Wiki 百科](https://zh.wikipedia.org/wiki/ICalendar)

- 墙内请看 [百度百科](https://baike.baidu.com/link?url=CNXZUdK4xnc-CCnlnwDgpxSZBvZaMaEQ3KkOlxndmvTEIpQ5kyichBHqcOEj8yUMB4MLC7JsH7hFs6b-Biy0rEYqV5GRH0dQkK0I8MriGy7)

     


- [这里](https://www.jianshu.com/p/8f8572292c58) 有关于 iCalendar 的语法说明

- [这里](https://icalendar.org/) 是 iCalendar 组织的主页

     

### 我的设备如何导入 iCalendar

#### `Windows 10`

双击打开`.ics`文件，即自动导入 Outlook 日历。

#### `macOS`

参阅 [这里](https://support.apple.com/zh-cn/guide/calendar/icl1023/mac)。

#### `iOS`和`Andriod`

[一个思路](https://zhuanlan.zhihu.com/p/35300266) 提供参考。

   

## 开始课程表抓取（xls 版本）

### 使用说明

1. [点击这里](https://raw.githubusercontent.com/smilonely/ClassCatch/master/classedit_xls.exe)，下载`classedit_xls.exe`文件
2. 将 `.xls `课程表更名为`1.xls`，放在与同目录下
3. 运行`classedit_xls.exe`，按提示输入信息
4. 稍等片刻，不出意外的话`.ics`文件将出现在同目录下

现在你就可以将`.ics`文件导入你的日历了。

   

### 运行环境

`Windows 10`（已测试）

`Windows X`

不支持`macOS`、`OSX`及`Linux`

这两个系统的小伙伴动手能力强的可以拿源码试试

   

### 其他说明

- 目前适配 2018 学年下学年课程表格式，可能还有些许bug。

  欢迎给我反馈 bug 。

- 项目源代码包含在`classedit_xls.py`中
- 项目所使用的模块：

```
import xlrd
import time
import datetime
import random
import string
import codecs
```

   

   

欢迎与我友好地讨论交流~