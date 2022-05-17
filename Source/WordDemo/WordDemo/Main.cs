using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordDemo
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.Title = "Word2003文件";
            sfd.FileName = "";
            sfd.Filter = "Word2003文件(*.doc)|*.doc";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //实例化COM
                Microsoft.Office.Interop.Word.ApplicationClass wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
                object fileobj = sfd.FileName;
                object nullobj = System.Reflection.Missing.Value;
                //打开指定文件（不同版本的COM参数个数有差异，一般而言除第一个外都用nullobj就行了）
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref fileobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj
                    );
                //取得doc文件中的文本
                string outText = doc.Content.Text;

                MessageBox.Show(ReadWord(doc, 1, 1, 1));
                //关闭文件
                doc.Close(ref nullobj, ref nullobj, ref nullobj);
                //关闭COM
                wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);
                //返回
                //MessageBox.Show(outText);

            }
        }

        /// <summary>
        /// 返回指定单元格中的数据
        /// </summary>
        /// <param name="tableIndex">表格号</param>
        /// <param name="rowIndex">行号</param>
        /// <param name="colIndex">列号</param>
        /// <returns>单元格中的数据</returns>
        public string ReadWord(Document doc, int tableIndex, int rowIndex, int colIndex)
        {
            //Give the value to the tow Int32 params.

            try
            {
                var table = doc.Tables[tableIndex];
                string text = table.Cell(rowIndex, colIndex).Range.Text.ToString();
                text = text.Substring(0, text.Length - 2);    //去除尾部的mark 
                return text;
            }
            catch
            {
                return "Error";
            }
        }
    }
}
