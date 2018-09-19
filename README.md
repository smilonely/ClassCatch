# QCU 课程表抓取

*目前仅更新了xls课程表抓取脚本，抓取网页的版本何时上传还是个未知数 ：）*

## xls课程表抓取

### 使用说明

1. [点击这里](https://raw.githubusercontent.com/smilonely/ClassCatch/master/classedit_xls.exe)，下载`classedit_xls.exe`文件
2. 将xls课程表更名为`1.xls`，放在与同目录下
3. 运行`classedit_xls.exe`，按提示输入信息
4. 稍等片刻，不出意外的话ics文件将出现在同目录下

现在你就可以将ics导入你的日历了。

------

目前适配2018学年下学年课程表格式，可能还有些许bug。

欢迎给我反馈bug。

------

运行环境：

`Windows 10`（已测试）

`Windows X`

不支持`macOS`、`OSX`及`Linux`

这两个系统的小伙伴动手能力强的可以拿源码试试

------



### 其他说明

- 项目源代码包含在`classedit_xls.py`中
- 项目所使用的模块：

```
import xlrd
import time
import datetime
import random
import string
import re
import os
import codecs
```



欢迎与我友好地讨论交流~