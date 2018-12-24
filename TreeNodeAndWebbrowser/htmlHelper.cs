using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TreeNodeAndWebbrowser
{
    public class htmlHelper
    {
        /// <summary>
        /// 替换模板中的占位符并由string的html生成html文件，返回文件名（路径+名）
        /// </summary>
        /// <param name="content">html模板</param>
        /// <param name="html">占位符处应该有的内容，动态生成的</param>
        /// <returns></returns>
        public static string GetHtmlInsteadPlaceholder(string content, string html)
        {
            //替换模板中的占位符
            content = content.Replace("$content$", html);
            //生成临时html文件
            var randomFilename = Guid.NewGuid().ToString();
            var folder = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "ReportsTemp";
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            var filename = folder + @"\" + randomFilename + ".html";
            HtmlToFile(filename, content);
            return filename;
        }
        /// <summary>
        /// html以string样子生成html格式的文件
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="content"></param>
        public static void HtmlToFile(string filename, string content)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.Write(content);
                    sw.Flush();
                }
            }
        }
        /// <summary>
        /// 读取html模板内容，返回string类型的html
        /// </summary>
        /// <returns></returns>
        public static string GetHtmlTemplate(string path)
        {

            return File.ReadAllText(path);//读取html文件
        }
    }
}
