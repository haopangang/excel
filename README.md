# JAVA解析Excel工具easyexcel

Java解析、生成Excel比较有名的框架有Apache poi、jxl。但他们都存在一个严重的问题就是非常的耗内存，poi有一套SAX模式的API可以一定程度的解决一些内存溢出的问题，但POI还是有一些缺陷，比如07版Excel解压缩以及解压后存储都是在内存中完成的，内存消耗依然很大。easyexcel重写了poi对07版Excel的解析，能够原本一个3M的excel用POI sax依然需要100M左右内存降低到KB级别，并且再大的excel不会出现内存溢出，03版依赖POI的sax模式。在上层做了模型转换的封装，让使用者更加简单方便

# 相关文档

* [关于软件](/abouteasyexcel.md)
* [快速使用](/quickstart.md)
* [常见问题](/problem.md)
* [更新记事](/update.md)





# 基于easyExcel版本
## VERSION : 1.2.4

# 原作者
姬朋飞（玉霄）

