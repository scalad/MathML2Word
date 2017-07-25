### 如何将MATHML转为Word文档 ###

[MATHML（Mathematical Markup Language，MathML）](https://zh.wikipedia.org/wiki/%E6%95%B0%E5%AD%A6%E7%BD%AE%E6%A0%87%E8%AF%AD%E8%A8%80)是一种基于XML的标准，用来描述数学符号和公式。它的目标是把数学公式集成到万维网和其他文档中。从2015年开始，MathML成为了HTML5的一部分和ISO标准。

由于数学符号和公式的结构复杂且符号与符号之间存在多种逻辑关系，MathML的格式十分繁琐。因此，大多数人都不会去手写MathML，而是利用其它的工具来编写，其中包括TeX到MathML的转换器。

有些时候，我们想要将MATHML导出到Word中方便查看，我们该怎样实现呢？这个时候我们还需要了解一下微软Office的[OMML(Office math markup language)]()标记语言，它是一种在WORD里面进行公式表达的标记语法，是以XML结构来存储的。遗憾的是，MATHML并不能直接转换为Word文档，它需要先转换为OMML。

那么如何将MathML转换为OMML？答案是使用一个转换文件——MML2OMML.xsl，这个文件是office自带的，位于目录：%ProgramFiles%\Microsoft Office\Office12\之下（若你用的是office 2016，则在%ProgramFiles%\Microsoft Office\Office16\目录）。

不知你是否知道，将MathML公式以文本的形式粘贴到Word中时，它会自动变成Word公式，这个操作的背后就是MML2OMML.xsl在起作用。同样的目录下还有一个文件OMML2MML.xsl，它的作用是反过来转换，我们这里用不到。

![](https://github.com/scalad/MathML2Word/blob/master/doc/image/20170725164449.png)
这个式子的MATHML的XML代码如下：

```xml
<math xmlns="http://www.w3.org/1998/Math/MathML" mathvariant='italic' display='inline'>
    <msub>
        <mi>R</mi>
        <mi>i</mi>
    </msub>
    <msup>
        <mtext></mtext>
        <mi>j</mi>
    </msup>
    <msub>
        <mtext></mtext>
        <mtext>kl</mtext>
    </msub>
    <mo>=</mo>
    <msup>
        <mi>g</mi>
        <mtext>jm</mtext>
    </msup>
    <msub>
        <mi>R</mi>
        <mtext>imkl</mtext>
    </msub>
    <mo>+</mo>
    <msqrt>
        <mn>1</mn>
        <mo>-</mo>
        <msup>
            <mi>g</mi>
            <mtext>jm</mtext>
        </msup>
        <msub>
            <mi>R</mi>
            <mtext>mikl</mtext>
        </msub>
    </msqrt>
</math>

```

具体的实现过程参考了[https://stackoverflow.com/questions/10993621/openxml-sdk-and-mathml](https://stackoverflow.com/questions/10993621/openxml-sdk-and-mathml)

