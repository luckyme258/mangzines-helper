1.将666.docx（文章） 666.xlsx（单词表格）转化为html格式。

- A4样式，左右对照，左侧文章，右侧生词的释义、音标、同义词。并且加入高亮黄色，以及颜色区分。

- cn-eng  优化版本

- 为了A4页面样式，排版美观，所以添加了强制换页
@@page  开始新的一页
@@pe     结束这一页


输出otuput.html



2. picture.py

- 用于将html得到的pdf文件 **output.pdf** ，变成高清的图片，调节dpi,600-800即可。
- 如何获得output.pdf文件？使用浏览器打开html文件，打印即可（搜索：如何用浏览器打印成pdf模式）


- pdf-maker.py 最初版本，基本没用了。

3. 制作效果

- 例子1

https://luckyme258.github.io/%E9%9B%85%E6%80%9D/Why%20Did%20a%20$10%20Billion%20Startup%20Let%20Me%20Vibe-Code%20for%20Them%E2%80%94and%20Why%20Did%20I%20Love%20It.html


- 例子2

https://luckyme258.github.io/%E5%9B%9B%E7%BA%A7/%E8%8B%B1%E8%AF%AD%E5%9B%9B%E7%BA%A7-1-%E7%BD%91%E7%BB%9C%E9%80%9F%E5%BA%A6.html

- 总目录

https://luckyme258.github.io/

4. 补充脚本 find-useful-pages.py


该脚本可自动读取同级`output.html`文件，识别并以“连续页码用-、非连续用、分隔”的格式，输出其中“右侧区域含单词卡片”的页码，无需人工逐页查看。
