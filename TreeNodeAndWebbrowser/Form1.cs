using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using static TreeNodeAndWebbrowser.Extension;

namespace TreeNodeAndWebbrowser
{
    public partial class Form1 : Form
    {
        Timer timer1;
        bool node_btn_fresh = false;
        public Form1()
        {
            InitializeComponent();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            //var node = new TreeNode("测试");
            //treeView1.Nodes.Add(node);
            //treeView1.AddTrizNode(null, new TreeNode("dfdferer"), Update, Remove, GetData()[0]);
            //treeView1.ExpandAll();

            node_btn_fresh = false;
            var p = new Person();
            var frm = new Form2(p);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                var node = new TreeNode("", 1, 2);
                treeView1.AddTreeNode(null, node, Update, Remove, p);
            }
            node_btn_fresh = true;
        }

        #region 编辑删除按钮的方法
        private void Update(TreeNode currentNode)
        {
            var p = ((TreeNodeTag)currentNode.Tag ).person;
            var frm = new Form2(p);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                currentNode.ResetNodeText(p.Name,true);//重绘text和编辑删除按钮的位置
            }
        }

        private void Remove(TreeNode currentNode)
        {
               // treeView1.Nodes.Remove(currentNode);
           
        }
        #endregion
        #region 测试数据

        private List<Person> GetData()
        {
            var datas = new List<Person>();
            datas.Add(new Person {
                Name="zhangsan",
                Age=20,
                Gender="male",
                Tel="12345678",
                Desc="asjdhfjas圣诞节发顺丰手动阀手动阀"
            });
            datas.Add(new Person {
                Name = "lisi",
                Age = 22,
                Gender = "male",
                Tel = "12334578",
                Desc = "asj恢复解放军s圣诞节发顺丰手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });

            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });

            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            datas.Add(new Person
            {
                Name = "amy",
                Age = 27,
                Gender = "female",
                Tel = "6553454578",
                Desc = "依然有人手动阀手动阀"
            });
            return datas;
        }
        private void GetSetData()
        {
            var datas = GetData();
            foreach (var a in datas)
            {
                var node = new TreeNode("", 1, 2);
                treeView1.AddTreeNode(null, node, Update, Remove, a);
            }
        }
        #endregion
        /// <summary>
        /// 生成测试1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // GenerateHTMLFile(new List<Person>());
            var lists = new List<Person>();
            foreach(TreeNode item in treeView1.Nodes)
            {
                var tag = item.Tag as TreeNodeTag;
                lists.Add(tag.person);
            }
            webBrowser1.Navigate(GenerateHTMLFile(lists));
        }
        /// <summary>
        /// 生成测试2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            var lists = new List<Person>();
            foreach (TreeNode item in treeView1.Nodes)
            {
                var tag = item.Tag as TreeNodeTag;
                lists.Add(tag.person);
            }
            webBrowser2.Navigate(GenerateHTMLFile(lists));
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            node_btn_fresh = true;
            treeView1.AfterExpand += TreeView1_After;//树形展开或者折叠都要重绘按钮
            treeView1.AfterCollapse += TreeView1_After;
            ResetNodeBtn();
            //测试数据
#if DEBUG
            GetSetData();
#endif
            webBrowser1.Navigate("about:blank");
            //while (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
            //{
            //    Application.DoEvents();
            //}
            //var a = Application.StartupPath + @"\htmlTempalte.html";
            //webBrowser1.Navigate(a);

            //webBrowser1.Document.Write();

            //this.webBrowser1.Url = new Uri("http://www.baidu.com");//指定url地址为百度首页
        }

        

        /// <summary>
        /// 树形展开或者折叠都要重绘按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TreeView1_After(object sender, TreeViewEventArgs e)
        {
            node_btn_fresh = false;
            treeView1.Nodes.ResetPictureBoxButtonVisibility();
            node_btn_fresh = true;
        }
        /// <summary>
        /// 树形滚动条滚动让按钮重绘
        /// </summary>
        private void ResetNodeBtn()
        {
            #region 树形滚动条滚动让按钮重绘
            timer1 = new Timer();
            timer1.Interval = 50;
            int lastNodeY = 0;
            timer1.Tick += (s1, e1) =>
            {
                if (node_btn_fresh && treeView1.Nodes.Count > 0)
                {
                    var curNodeY = treeView1.Nodes[treeView1.Nodes.Count - 1].Bounds.Location.Y;
                    if (curNodeY != lastNodeY)
                    {
                        lastNodeY = curNodeY;
                        treeView1.Nodes.ResetPictureBoxButtonVisibility();
                    }
                }
            };
            timer1.Start();
            #endregion
        }
        /// <summary>
        /// 动态生成html文件
        /// </summary>
        /// <param name="lists"></param>
        /// <returns></returns>
        public string GenerateHTMLFile(List<Person> lists)
        {
            string path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            path += @"\htmlTemplate.html";
            string content = htmlHelper.GetHtmlTemplate(path);
            //动态生成html
            string html = GetHtml(lists);

            return htmlHelper.GetHtmlInsteadPlaceholder(content, html);
        }
        /// <summary>
        /// 获取html动态部分
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private string GetHtml(List<Person> lists)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in lists)
            {
                sb.Append("<tr>");
                sb.Append("<td>");
                sb.Append(item.Name);
                sb.Append("</td>");
                sb.Append("<td>");
                sb.Append(item.Tel);
                sb.Append("</td>");
                sb.Append("<td>");
                sb.Append(item.Desc);
                sb.Append("</td>");
                sb.Append("</tr>");
            }
            return sb.ToString();
        }

        private void btn_html_Click(object sender, EventArgs e)
        {
            //var wb = new WebBrowser();            
            //wb.DocumentText = string.Empty; //赋个值，否则Document为null会报错
           // wb.Document.OpenNew(true);
           //if(webBrowser1.DocumentText.EndsWith("\0")&& webBrowser2.DocumentText.EndsWith("\0"))
           // {
           //     MessageBox.Show("无html");
           //     return;
           // }
            var htmlList = new List<HtmlPendingToExcel>();
            foreach (TabPage tp in tabControl1.TabPages)
            {
                foreach (Control c in tp.Controls)
                {
                    if (c is WebBrowser)
                    {
                        var w = c as WebBrowser;
                        if (!string.IsNullOrEmpty(w.DocumentText))
                        {
                            htmlList.Add(new HtmlPendingToExcel
                            {
                                Html = w.DocumentText,
                                Name = tp.Text
                            });
                        }
                    }
                }
            }
            StringBuilder s = new StringBuilder();
            for (int i = 0; i < htmlList.Count; i++)
            {
                var html = htmlList[i];
                if (i == 0)
                {
                    s.AppendLine("<h3 style='text-align: center;'>" + html.Name + "</h3>");
                }
                else
                {
                    s.AppendLine("<h3 style='page-break-before:always;text-align: center;'>" + html.Name + "</h3>");//第二个html放到下一页
                }
                s.Append(html.Html);
            }
            var sfd = new SaveFileDialog
            {
                Title = "请选择htmll的保存位置",
                FileName = "html",
                Filter = "*.html|*.html"
            };
            var dr = sfd.ShowDialog(this);
            string path = string.Empty;
            if (dr != DialogResult.OK)
            {
                return;
            }
            else
            {
                path = sfd.FileName;
            }
            htmlHelper.HtmlToFile(path,s.ToString());
        }

        private void btn_excel_Click(object sender, EventArgs e)
        {
            var htmlList = new List<HtmlPendingToExcel>();
            foreach (TabPage tp in tabControl1.TabPages)
            {
                foreach (Control c in tp.Controls)
                {
                    if (c is WebBrowser)
                    {
                        var wb = c as WebBrowser;
                        HtmlElementCollection tables = wb.Document.GetElementsByTagName("table");
                        foreach (HtmlElement tb in tables)
                        {
                            htmlList.Add(new HtmlPendingToExcel
                            {
                                Name = tp.Text,
                                Html = tb.OuterHtml,
                                RowCount = tb.GetElementsByTagName("TR").Count
                            });
                        }
                    }
                }
            }


            var tabName = tabControl1.SelectedTab.Text;
            var sfd = new SaveFileDialog
            {
                Title = "请选择Excel的保存位置",
                FileName = tabName,
                Filter = "(*.xlsx)|*.xlsx"
            };
            var dr = sfd.ShowDialog(this);
            string path = string.Empty;
            if (dr != DialogResult.OK)
            {
                return;
            }
            else
            {
                path = sfd.FileName;
            }
            ApplicationClass xlApp = new ApplicationClass();
            try
            {
                //Excel.ApplicationClass xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;
                xlApp.Visible = false;

                Workbooks workbooks = xlApp.Workbooks;
                Workbook workBook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                while (workBook.Worksheets.Count > 1)
                {
                    workBook.Worksheets.Delete();
                }

                var obj = Clipboard.GetDataObject(); bool isFirstSheet = true;
                for (int m = htmlList.Count - 1; m >= 0; m--)
                {
                    var html = htmlList[m];

                    if (isFirstSheet) { isFirstSheet = false; }
                    else { workBook.Worksheets.Add(); }

                    Worksheet workSheet = (Worksheet)workBook.Worksheets[1];

                    workSheet.Name = html.Name;

                    Clipboard.SetText(html.Html, TextDataFormat.Text);

                    workSheet.PasteSpecial();

                    for (int i = 1; i < html.RowCount; i++)
                    {
                        Range cell = workSheet.get_Range("A" + i, "A" + i);
                        if (cell.Value == null) { continue; };
                        cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        string str = cell.Value.ToString().Trim();
                        if (str.StartsWith("-"))//(开头的单元格在excel中变成-
                        {
                            cell.NumberFormatLocal = "@";
                            cell.Value = string.Format("({0})", str.Substring(1));
                        }
                    }

                    workSheet.Columns.AutoFit();
                }
                Clipboard.SetDataObject(obj);
                workBook.SaveAs(path);
                //workbooks.Add(true).SaveAs(path);

            }
            catch (Exception ex)
            {

            }
            finally
            {
                try
                {
                    //xlApp.Workbooks.Close();
                    //xlApp.Quit();  //自动打开Excel，抛异常
                }
                catch { }
            }
        }

        private void btn_preview_Click(object sender, EventArgs e)
        {

        }
    }

}
