1、excel的上传模板开发

2、httpServletResponse的属性设置   https://blog.csdn.net/qq_34666857/article/details/104838171
响应消息头名称         说明
Server             一种标明web服务器软件几版本号的头标
Content-Type       返回文档时所采用的的MIME类型
Transfer-Encoding  太长了
Date               发送HTTP的日期
Content-Encoding   用于说明文本在传输过程中已采用的编码方式
Content-Length     响应内容的长度，以字节为单位
Expire             设定一段时间，时间到文档过期，不应再继续缓存
Refresh            多少秒后浏览器重新载入页面
Cache-Control      指定缓存机制，若取no-cache值表示阻止浏览器缓存页面
Location           浏览器重新连接到的URL
Content-Disposition 通过这个报头，可以让浏览器询问用户将响应存储到磁盘上给定名称的文件中  attachment; filename=\"" + fileName + "\"
Set-Cookie         浏览器记下cookie，推荐使用addCookie

3、常用的MIME类型如下：

　　超文本标记语言文本 .html,.html text/html
　　普通文本 .txt text/plain
　　RTF文本 .rtf application/rtf
　　GIF图形 .gif image/gif
　　JPEG图形 .jpeg,.jpg image/jpeg
　　au声音文件 .au audio/basic
　　MIDI音乐文件 mid,.midi audio/midi,audio/x-midi
　　RealAudio音乐文件 .ra, .ram audio/x-pn-realaudio
　　MPEG文件 .mpg,.mpeg video/mpeg
　　AVI文件 .avi video/x-msvideo
　　GZIP文件 .gz application/x-gzip
　　压缩文件.rar application/octet-stream
　　压缩文件.zip application/x-zip-compressed
　　TAR文件 .tar application/x-tar