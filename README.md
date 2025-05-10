# ConfigTools
配置表生成复制工具

---

这是一个excel表转json/xml/lua的工具。

PS：仅支持utf8的excel表

---
### 表格式

第1行是字段，原则上不允许除数字和整数之外的字符，除了字段第1个字符的特殊含义：
1. 不生成该字段(#)
2. 仅客户端生成(!)
3. 仅服务端($)

第2行是数据类型，支持以下的类型：
1. bool
2. int
3. float
4. json

第3行是注释(生成时不会导出)

第4~N行是数据，其中第1列是主键，最好是int或string类型，避免可能存在的问题。

具体的例子可参考excel/Test.xlsx

### config.json字段说明
---

input：放置excel的文件夹目录

output：生成json/xml/lua配置的临时目录（会自动在output指定的位置添加client和server目录）

move：可以指定client/server和对应的格式，自动复制到对应的目录。（可缺省参数）

具体的例子可参考config.json

---

### 运行方式
双击run.bat或runOnce.bat

其中，run.bat是监测文件变化自动生成，runOnce.bat立刻运行1次。

PS：将读取上一层目录的config.json的配置。

---

### 例子运行方式
双击example.bat或exampleOnce.bat

PS：将读取本目录的config.json的配置。