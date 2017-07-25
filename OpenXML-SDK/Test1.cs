using System;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;

public class Test1
{
    public XslCompiledTransform XslTransforms;
    XslTransforms = new XslCompiledTransform();
    XslTransforms.Load(@"D:\DocEditor\Microsoft Office\Office14\MML2OMML.xsl");


    void CreateWord()
    {
        object path;//文件路径
        string strContent;//文件内容
        MSWord.Application wordApp;//Word应用程序变量
        MSWord.Document wordDoc;//Word文档变量
        path = "d:\\myWord.doc";//保存为Word2003文档
        // path = "d:\\myWord.doc";//保存为Word2007文档
        wordApp = new MSWord.ApplicationClass();//初始化
        if (File.Exists((string)path))
        {
            File.Delete((string)path);
        }
        //由于使用的是COM 库，因此有许多变量需要用Missing.Value 代替
        Object Nothing = Missing.Value;
        //新建一个word对象
        wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        //WdSaveDocument为Word2003文档的保存格式(文档后缀.doc)\wdFormatDocumentDefault为Word2007的保存格式(文档后缀.docx)
        object format = MSWord.WdSaveFormat.wdFormatDocument;
        //将wordDoc 文档对象的内容保存为DOC 文档,并保存到path指定的路径
        wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        //关闭wordDoc文档
        wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
        //关闭wordApp组件对象
        wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        Response.Write("<script>alert('" + path + ": Word文档创建完毕!');</script>");
    }

    static void Main(string[] args)
    {

    }
}
