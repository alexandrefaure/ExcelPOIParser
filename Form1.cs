using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using NPOI.XWPF.UserModel;

namespace TestExcelParser
{
    public partial class Form1 : Form
    {
        private static string excelFilePath = "D:\\tests\\4573-DCE-DPGF ind B.xlsx";
        private static string wordFilePath = "C:\\Users\\FAURE\\Documents\\die_4816176\\mail N2.docx";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var path = GetFilePath(excelFilePath);

            var parser = new ExcelParser();
            var resultNodeCollection = parser.Parse(path);

            FillListView(resultNodeCollection);
        }

        private static string GetFilePath(string filePath)
        {
            var path = filePath;
            if (string.IsNullOrEmpty(filePath))
            {
                var fileManager = new FileManager();
                fileManager.Open();

                path = fileManager._filePath;
            }

            return path;
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

        private void button2_Click(object sender, EventArgs e)
        {
            var path = GetFilePath(wordFilePath);

            XWPFDocument document = null;
            try
            {
                using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    document = new XWPFDocument(file);
                }
            }
            catch (Exception ex)
            {
                throw new System.InvalidOperationException(ex.Message);
            }
        }
    }
}