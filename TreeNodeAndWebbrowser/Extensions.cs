using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TreeNodeAndWebbrowser
{
    static class Extensions
    {
        public static void AddTreeNode(TreeNode parentNode,Person person)
        {
            TreeNode currentNode = new TreeNode(person.Name);
            // currentNode.Text = person.Name;
            var tag = new TreeNodeTag();
            tag.pictureBox_Edit = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.CenterImage,//图片中间显示
                Cursor = Cursors.Hand,//光标形状
                Size = new System.Drawing.Size(16, 16),//大小16*16
                Image = Properties.Resources.update,//图片
                BackColor = System.Drawing.Color.Transparent,
                Location=currentNode.Bounds.Location,
                Tag = currentNode,
                Visible = false
            };
            tag.person = person;
            currentNode.Tag = tag;
            parentNode.Nodes.Add(currentNode);
        }
    }

   
    public class TreeNodeTag
    {
        public PictureBox pictureBox_Edit;
        public PictureBox pictureBox_Delete;
        public Person person;
        public TreeNodeTag()
        {
            pictureBox_Delete = new PictureBox();
            pictureBox_Edit = new PictureBox();
        }
    }
}
