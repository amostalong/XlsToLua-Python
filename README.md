# XlsToLua-Python

把xls转为lua脚本的python工具:

1.首先要按照python3.6 + : https://www.python.org/downloads/

2.安装的时候安装好自带的pip，并配置好path

3.windows使用bat可以自己生产，生产的结果会存储在当前目录下的一个以时间为名字的文件夹中

4.多sheet处理: 支持一个xls中的多个sheet, 每一个目标sheet都会被生成为一个table

5.excel格式：

A列用来做ID，必须是string或者int格式。必须包含一个data 的 sheet。
第一行是只做注释用，第二行是key得名字，第三行是value得类型，第五行开始才是每一列得值。

值类型：

`int`: 整数。

`string`: 文本。

`float`: 小数。

`bool`： 布尔值

array[int,string,float,bool]:数组,会自动生成 {[1] = x, [2] = y, [3] = z} 这种格式。

table:自由格式，完全由策划自己去配置，会直接复制到lua配置表中
    这个格式基本没有错误检测功能
    文本必须用‘’来表示：比如 [1] = A，必须写成 [1] = 'A'
    lua table通常格式为:
        
    Param2 = {
        
        A = 100,

        B = {

            [1] = 'A',

            [2] = 2,

        }
    },

在excel 的 单元格中为： { A = 100, B = { [1] = 'A', [2] = 2 } }

emmylua注释： 这个python脚本也会生产对应的emmylua格式的类型说明（table不会生产，没有去解析table内部，可能以后会去弄）

参数:

1. `-sh:` 必须参数后面接目标sheet的名字
2. `-fd:` 可选,当使用时会用生成时间建立一个新的文件夹

例如:`python XlsToLua.py -sh data test -fd`
表明目标: sheet = ['data','test'],folder = true