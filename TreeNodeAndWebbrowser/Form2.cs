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
    public partial class Form2 : Form
    {
        Person _p;
        
        public Form2(Person p)
        {
            _p = p;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _p.Name = textBox1.Text;
            _p.Tel = textBox2.Text;
            _p.Desc = textBox3.Text;
            this.DialogResult = DialogResult.OK;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            textBox1.Text = _p.Name;
            textBox2.Text = _p.Tel;
            textBox3.Text = _p.Desc;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
