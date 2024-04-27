### 是什么
这是一个用于excel填表的工具，把input表中内容填到output.xlsx中。
只是用了openpyxl读取写入文件，然后用tkinter包装了下
跟复制粘贴相比有什么优势？
- 多方件处理时更方便
- 当单元格值在分散的多个位置时，这个更方便
### tips
- python 环境3.8以上
- excel文件格式为xxx.xls,而不是xxx.xlsx,因为发现很多地方导出内容还是用xls格式的，如果你是用xxx.xlsx，有二种方法，
    1.去掉代码中转换excel语句,(所以你的设备上需要装有excel程序) 
    2.打开这个excel文件另存为xls格式
- 这个不能处理合并的单元格
- 这个不能处理一个目录下的xlsx文件
- 默认读取excel第一张表(sheet1)
- 读取时有保留excel表原格式
### 使用方法
#### 我只是使用这个程序
- 直接下载释放的版本，在这里
#### 我要看看这个程序
- git clone xxx
- cd xxx
- python -m venv myenv #建虚拟环境
- /myenv/scripts/activate #激活虚拟环境
- pip install -r env.txt #安装依赖
- python main.py #运行程序
### 程序使用
- 文件解释,input.xls是输入的源数据，template.xlsx是中间模板，output.xlsx是最后生成的表。
- 先运行一次main.exe,在当前目录会生成input.txt文件,
- 编辑data.csv,对比input.txt文件和template.xlsx文件,其中:
    - 第一列是:要填入的数据在templates.xlsx的单元格位置，如A2,B3,我希望你能意识到，当单元格分散时，用这个repo能更快处理数据
    - 第二列是:要填入的数据在input.txt中的行数;
- 点击读取excel，完成excel填表

### 补充点
- 在Excel中，"A1"这种表示单元格的方法称为列-行表示法或RC表示法（Row-Column Notation）。这种表示法通过组合列标和行号来唯一标识工作表上的每一个单元格。
- 具体解释如下：
- 列标：列标是指单元格所在的列，使用英文字母表示。Excel中的列标从A开始，依次为B、C、...，一直到Z，然后继续使用双字母组合如AA、AB、...，一直到AZ，接着是BA、BB等，以此类推。
- 行号：行号是指单元格所在的行，使用数字表示。Excel中的行号从1开始，依次为2、3、...，一直到1048576，这是Excel工作表的最大行数。
- 因此，当我们使用"A1"时，我们指的是位于第一列（Column A）和第一行（Row 1）的单元格。同理，"B2"指的是第二列第三行的单元格，"C3"指的是第三列第三行的单元格。



