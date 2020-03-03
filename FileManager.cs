using System.IO;
using System.Windows.Forms;

namespace TestExcelParser
{
    public class FileManager
    {
        private string _fileContent = string.Empty;
        public string _filePath = string.Empty;

        public void Open()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    _filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (var reader = new StreamReader(fileStream))
                    {
                        _fileContent = reader.ReadToEnd();
                    }
                }
            }

            MessageBox.Show(_fileContent, "File Content at path: " + _filePath, MessageBoxButtons.OK);
        }
    }
}