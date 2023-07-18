# 表单提取工具

## 1. 介绍
该工具用于提取excel表单中的数据， 使用时需要声明输出表单的列名，以及对应输入表单的行号和列号。
比如
```csv
输出列名, 行号, 列号
(单据头)项目号, 1, I
(单据头)介质, 14, I
```
其中 列号是excel里的字母列号，行号是excel里的数字行号，从1开始计数。

## 2. 安装
安装对应依赖

使用前确保安装了python3.6以上版本，以及pip3
#### 2.1.1 Windows
1. 安装python3
打开 WEB 浏览器访问 https://www.python.org/downloads/windows/ ，
一般就下载 executable installer，x86 表示是 32 位机子的，x86-64 表示 64 位机子的。

2. 安装 pip3
下载[get-pip.py](https://bootstrap.pypa.io/get-pip.py)到本地，然后执行

### 2.1.2 Mac
1. 安装 python3
如果没有 python3，则先安装：
```bash
brew install python3
```
2. 安装 pip3：

```bash
curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
python3 get-pip.py
```

### 2.2 检查版本
```bash
python3 -V
pip3 -V
```


### 2.3 安装依赖
```bash
pip3 install -r requirements.txt
```

## 3. 使用
### 3.1 命令行

#### 3.1.1.指定参数的方式
```bash
python3 main.py -i/--input_file <input_file> -o/--output_file <output_file> -c/--config_file <config_file> -s/--sheet_name <sheet_name>
```

可以只提供部分参数，未提供的参数将使用默认值

#### 3.1.2. 使用默认参数的方式
```bash
python3 main.py
```



### 3.2 参数说明
| 参数 | 说明 |
| --- | --- |
| input_file | 输入文件路径 |
| output_file | 输出文件路径 |
| config_file | 配置文件路径 |
| sheet_name | 表单名称 |

注意： 
1. 输出文件路径的文件夹必须存在，否则会报错， 输入文件路径和配置文件路径必须存在，否则会报错
2. 输入/输出文件必须为xlsx格式，否则会报错
3. 配置文件必须为csv格式，否则会报错
4. 如果输出文件已存在，需要先删除，否则会报错

### 3.3 配置文件说明
配置文件为csv格式，第一行为列名，第二行开始为数据，每一行为一个配置项，配置项包括三个字段，分别为输出列名、行号、列号，以逗号分隔，如下所示：
```csv
输出列名, 行号, 列号
(单据头)项目号, 1, I
(单据头)介质, 14, I
```
## 4. 打包
```bash
pyinstaller -F main.py
```