从CzBix/lzybetter大神那里forked过来的。动机是lzybetter版本的code模式失效，微软改了新的code模式，于是我将其改为手动选取code。
针对之前版本，新增了对附件的收集功能，具体运行attachments.py

# WizNote-to-OneNote
笔记映射关系如下： 

| 为知笔记             | OneNote          |
| -------------------- | ---------------- |
| 第一级文件夹         | 笔记本           |
| 第二级、第三级文件夹 | 分页<sup>1</sup> |
| 第四级文件夹到笔记   | 页面<sup>2</sup> |

1如果只存在一级文件夹，分页的名字将和笔记本相同；若只有两级文件夹，则分页名为第二级文件夹名；若存在三级文件夹，则分页名为第二级文件夹名-第三级文件夹名

2页面名与笔记名相同

另外，增加了断点续传的功能，程序错误终止或者token有效期到了以后可以重新接着上次的内容传输。
断点续传会在当前文件夹中创建一个名为count.txt的文件，用于记录当前传输进程，请勿随意删除。

一个从为知笔记导入数据到 OneNote 的脚本

## 已知问题
1. OneNote 本身的限制。比如不支持自定义 CSS，文件大小等。（具体请看 [OneNote API](https://dev.onenote.com/docs)）
2. 不支持加密文档（请先在为知客户端手动解密）
3. 不支持附件（请在导入完后手动添加）
4. ~~不支持断点续传，且从微软获取的 token 有效期为 1 个小时。（意味着超过 1 小时会上传失败。如此巨量文档建议您还是不要轻易迁移知识库了。）~~
token还是一个小时的有效期，但是有断点续传了，token过期后只需要再次启动就可以了
5. Linux 环境测试通过（测试环境：Ubuntu 18.04 + Python 3.6.6）/ win11环境（py37）测试通过。
6. 对于单次上传笔记数量过多（>500）、单个笔记元素过多（>500）、微信收藏笔记不完整（有元素缺失）的情况可能导致进程停止。单次上传次数过多需要过一会儿重新启动程序。对于单个文件出问题的，需要手动迁移，并删除原笔记，包括清空回收站。所以建议提前备份原来的为知笔记目录。


## 操作步骤
首先下载脚本并安装依赖

1. 登陆为知客户端，等待全部文件完成同步之后退出客户端
2. 执行脚本，根据提示输入账户名或文件夹路径（Windows 上路径为 `我的文档\My Knowledge\Data\{account}\`）
3. 打开提供的登录链接进行授权，然后复制并输入跳转后的空白页链接，需要手动选取链接中的code，位于'https://login.live.com/oauth20_desktop\.srf\?code='和可能存在的'&lc=****'之间，长度约为45+字符。
4. ~~选择目标笔记本名称（请输入新的名字以避免冲突）~~
5. 开始导入，等待导入完成后确认内容正确，并根据需要~~手动上传附件~~。
6. 【新增】附件可以通过attachments.py，从为知笔记文件夹路径迁移到一个新的目录，自动过滤掉.ziw文件。可以通过level参数限制新目录的级别，支持0（全部附件放在同个目录下）、1（放在笔记的一级目录下，比如我的笔记/、微信收藏/、我的日记/），2以上以此类推
