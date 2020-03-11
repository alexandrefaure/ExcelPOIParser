using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using NPOI.XWPF.UserModel;
using TestExcelParser.Model;

namespace TestExcelParser
{
    public partial class Form1 : Form
    {
        private static string excelFilePath = @"D:\tests\Extractions NPOI\4573-DCE-DPGF ind B.xlsx";
        private static string wordFilePath = @"D:\tests\Extractions NPOI\B.4 Bordereau des prix.docx";

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
                    var tables = document.GetXWPFDocument().Tables;

                    foreach (var table in tables)
                    {
                        var content = table.Text;
                        var rows = table.Rows;

                        foreach (var row in rows)
                        {
                            var cells = row.GetTableCells();
                            foreach (var cell in cells)
                            {
                                var wordElement = new WordElement
                                {
                                    style = new Style
                                    {
                                        
                                    }
                                };
                                var text = cell.GetText();
                                var color = cell.GetColor();

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new System.InvalidOperationException(ex.Message);
            }
        }
    }
}