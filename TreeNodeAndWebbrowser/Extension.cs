
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace TreeNodeAndWebbrowser
{
    public delegate void TreeNodeButtonClickEvent(TreeNode currentNode);
    public static class Extension
    {
        public static string Note(this TreeNode tn)
        {
            if (string.IsNullOrEmpty(tn.ToolTipText))
            {
                return tn.Text;
            }
            else
            {
                return tn.ToolTipText;
            }
        }
        /// <summary>
        /// 为treeview控件增加树节点，同时在特定层级节点名称右侧附加“修改”和“删除”图片按钮
        /// </summary>
        /// <param name="tv">当前树型组件</param>
        /// <param name="parentNode">需要在哪个节点下增加节点，如果是root，则传入null即可</param>
        /// <param name="currentNode">需要增加的节点</param>
        /// <returns>The zero-based index value of the System.Windows.Forms.TreeNode added to the tree node collection.</returns>
        public static int AddTreeNode(this TreeView tv, TreeNode parentNode, TreeNode currentNode, TreeNodeButtonClickEvent editNodeMethod, TreeNodeButtonClickEvent removeNodeMethod, Person person)
        {
            int index = 0;
            //currentNode.Text = currentNode.Text.Substring(0, maxDisplayLength) + (currentNode.Text.Length > maxDisplayLength ? "..." : string.Empty);
            ResetNodeText(currentNode, person.Name);
            if (parentNode == null) { index = tv.Nodes.Add(currentNode); }
            else { index = parentNode.Nodes.Add(currentNode); }
            var tag = new TreeNodeTag();
            // tag.EntityType = entityType;
            currentNode.Tag = tag;
            if (editNodeMethod != null)
            {
                var picBtnEdit = new PictureBox
                {
                    Parent = tv,//父容器
                    Image = Properties.Resources.update,//图片
                    //为什么不将以下属性放在SetPictureBoxButton中设置：按钮一旦生成，以下属性是确定了的，不确定只是按钮位置
                    //SetPictureBoxButton只是重绘位置属性
                    BackColor = System.Drawing.Color.Transparent,//图片背景色透明
                    Tag = currentNode,//保存要操作的的对象节点
                    SizeMode = PictureBoxSizeMode.CenterImage,
                    Cursor = Cursors.Hand,
                    Size = new System.Drawing.Size(16, 16),
                };
                tag.PictureBox_Edit = picBtnEdit;
                tag.person = person;
                picBtnEdit.Click += (s, e) =>
                {
                    var curNode = (s as PictureBox).Tag as TreeNode;
                    editNodeMethod(curNode);
                };
            }
            if (removeNodeMethod != null)
            {
                var picBtnRemove = new PictureBox
                {
                    Parent = tv, //父容器              
                    Image = Properties.Resources.delete,//图片
                    BackColor = System.Drawing.Color.Transparent,//图片背景色透明
                    Tag = currentNode,//保存要操作的的对象节点
                    SizeMode = PictureBoxSizeMode.CenterImage,
                    Cursor = Cursors.Hand,
                    Size = new System.Drawing.Size(16, 16),
                };
                tag.PictureBox_Remove = picBtnRemove;
                picBtnRemove.Click += (s, e) =>
                {
                    var curNode = (s as PictureBox).Tag as TreeNode;
                    var confirm = MessageBox.Show(string.Format("确定删除{0}吗？", currentNode.Text), "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (confirm == DialogResult.OK)
                    {
                        removeNodeMethod(curNode);//执行删除操作，删掉节点
                        var tag1 = curNode.Tag as TreeNodeTag;
                        tag1.PictureBox_Edit.Dispose();//删掉编辑按钮
                        tag1.PictureBox_Remove.Dispose();//删掉删除按钮
                        curNode.ClearChildNodeButtons();//清除所有孩子节点的编辑和删除按钮
                        curNode.Remove();//移除当前节点
                        //tv.Update();//更新treeview
                        tv.Nodes.ResetPictureBoxButtonVisibility();//重绘全部编辑和删除按钮
                    }
                };
            }
            ResetTreeNode(currentNode, currentNode.IsVisible);
            return index;
        }
        /// <summary>
        /// 删除节点时清除子节点的编辑和删除按钮
        /// </summary>
        /// <param name="curNode"></param>
        public static void ClearChildNodeButtons(this TreeNode curNode)
        {
            foreach (TreeNode node in curNode.Nodes)
            {
                var childTag = node.Tag as TreeNodeTag;
                if (childTag.PictureBox_Edit != null)
                {
                    childTag.PictureBox_Edit.Dispose();
                }
                if (childTag.PictureBox_Remove != null)
                {
                    childTag.PictureBox_Remove.Dispose();
                }
            }
        }
        /// <summary>
        /// 计算并重绘图片按钮的位置
        /// </summary>
        /// <param name="nodes"></param>
        public static void ResetPictureBoxButtonVisibility(this TreeNodeCollection nodes)
        {
            foreach (TreeNode currentNode in nodes)
            {
                ResetTreeNode(currentNode, currentNode.IsVisible);
                currentNode.Nodes.ResetPictureBoxButtonVisibility();
            }
        }
        /// <summary>
        /// 计算图片按钮的位置
        /// </summary>
        /// <param name="currentNode"></param>
        /// <param name="Visible"></param>
        private static void ResetTreeNode(TreeNode currentNode, bool Visible = true)
        {
            TreeNodeTag tag = (TreeNodeTag)currentNode.Tag;

            int offset_edit = currentNode.Bounds.Width + 44;//编辑按钮的左偏移量   

            int offset_remove = offset_edit + (tag.PictureBox_Edit == null ? 0 : 22);//删除按钮的左偏移量   

            currentNode
                .SetPictureBoxButton(tag.PictureBox_Edit, offset_edit)
                .SetPictureBoxButton(tag.PictureBox_Remove, offset_remove);
        }
        /// <summary>
        /// 重绘图片按钮的位置
        /// 什么情况下需要重绘位置:删除、编辑后
        /// </summary>
        /// <param name="currentNode"></param>
        /// <param name="pb"></param>
        /// <param name="offset"></param>
        /// <param name="visible"></param>
        /// <returns></returns>
        private static TreeNode SetPictureBoxButton(this TreeNode currentNode, PictureBox pb, int offset, bool visible = true)
        {
            if (pb != null)
            {
                var B = pb;
                B.Location = currentNode.Bounds.Location;
                B.Top = currentNode.Bounds.Location.Y + 1;
                B.Left = offset;
                //B.Visible = visible;//默认是显示的
            }
            return currentNode;
        }

        /// <summary>
        /// 重绘treenode的text和ToolTipText以及编辑和删除按钮的位置
        /// </summary>
        /// <param name="node"></param>
        /// <param name="str"></param>
        /// <param name="isResetLocation">是否重绘编辑和删除按钮的位置</param>
        public static void ResetNodeText(this TreeNode node, string str, bool isResetLocation=false)
        {
            node.Text = str.Length > 11 ? str.Substring(0, 8) + "..." : str;
            node.ToolTipText = str.Length > 11 ? str : "";
            if(isResetLocation) ResetTreeNode(node);
        }
        
        public class TreeNodeTag
        {
            public PictureBox PictureBox_Edit { get; set; }
            public PictureBox PictureBox_Remove { get; set; }
            public Person person { get; set; }
            public TreeNodeTag()
            {
                //this.tag = new Pipeline();
            }
        }
        public class HtmlPendingToExcel
        {
            public string Name { get; set; }
            public int RowCount { get; set; }
            public string Html { get; set; }
        }
    }
}
