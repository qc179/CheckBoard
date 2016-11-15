# CheckBoard
### 运行环境
- python-3.5.2，用安装包安装即可。
- 首次运行前，先运行本目录中mod里的checkmod.py，将会安装必要模块，在你的电脑上只需要安装一次即可，若安装不成功可重复运行。


## checkboard.py
### 简要
- checkboard.py：检查配置在board表上的版面，不限类型
- url存在且is_active=1则判断为已配置，输出配置的bid，以及有无重复和重复的数量。
- 若url是贴吧，则会以fid=101为前提条件，吧名存在则判断已配置，输出贴吧的bid

### 方法
- 打开board.xls
- 从第二行起填入name和url（来自信息源）
- 保存关闭board.xls
- 双击运行checkboard.py
- 运行结束，查看输出文件“checkboard_result_xxxxxxxx_xxxxxx.xls”

### 提醒
- 信息源的name或url尽量不要包含空格、换行符，不要在board.xls中插入空白行，运行前先手动检查，避免重复操作。


## checknew.py
### 简要
- checknew.py：检查配置在newssite表上的版面，不限类型
- url存在且is_active=1则判断为已配置，输出配置的nsid，以及有无重复和重复的数量。

### 方法
- 打开news.xls
- 从第二行起填入name和url（来自信息源）
- 保存关闭news.xls
- 双击运行checknews.py
- 运行结束，查看输出文件“checknews_result_xxxxxxxx_xxxxxx.xls”

### 提醒
- 信息源的name或url尽量不要包含空格、换行符，不要在news.xls中插入空白行，运行前先手动检查，避免重复操作。