using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TreeNodeAndWebbrowser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var node = new TreeNode("测试");
            treeView1.Nodes.Add(node);
            Extensions.AddTreeNode(node, GetData()[0]);
            treeView1.ExpandAll();
        }
        #region
        //数据
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
            return datas;
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show((treeView1.Nodes[0].Nodes[0].Tag as TreeNodeTag).person.Desc);
            }
            catch { }
        }
    }
}
