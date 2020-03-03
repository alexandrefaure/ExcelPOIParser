using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TestExcelParser
{
    public partial class Form1 : Form
    {
        private static string filePath = "D:\\tests\\4573-DCE-DPGF ind B.xlsx";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                var fileManager = new FileManager();
                fileManager.Open();

                filePath = fileManager._filePath;
            }

            var parser = new ExcelParser();
            var resultNodeCollection = parser.Parse(filePath);

            FillListView(resultNodeCollection);
        }

        private void FillListView(TreeNodeCollection resultNodeCollection)
        {
            foreach (TreeNode treeNode in resultNodeCollection)
            {
                var treeNodeList = new List<TreeNode>();

                foreach (TreeNode treeSubNode in treeNode.Nodes)
                {
                    treeNodeList.Add(treeSubNode);
                }

                var treeNodeToAdd = new TreeNode(treeNode.Text, treeNodeList.ToArray());

                treeView.Nodes.Add(treeNodeToAdd);
            }
        }
    }
}