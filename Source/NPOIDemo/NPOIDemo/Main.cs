using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using NPOI.XWPF.UserModel;

namespace NPOIDemo
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
            sfd.Title = "Word2007文件";
            sfd.FileName = "";
            sfd.Filter = "Word2007文件(*.docx)|*.docx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                StringBuilder sb = new StringBuilder();
                using (FileStream stream = File.OpenRead(sfd.FileName))
                {
                    XWPFDocument doc = new XWPFDocument(stream);
                    var tables = doc.Tables;
                    foreach (var table in tables)    //遍历表格  
                    {
                        if (table.Rows.Count > 0)
                        {
                            var row0 = table.GetRow(0);
                            for (int i = 0; i < row0.GetTableCells().Count;i++)
                            {
                                var c0 = row0.GetCell(i);        //获得单元格0  
                                foreach (var para in c0.Paragraphs)
                                {
                                    string text = para.ParagraphText;
                                    //处理段落      
                                    sb.Append(text + ",");
                                }
                            }
                        }
                        sb.AppendLine("");
                        continue;
                        foreach (var row in table.Rows)    //遍历行  
                        {
                            var c0 = row.GetCell(0);        //获得单元格0  
                            foreach (var para in c0.Paragraphs)
                            {
                                string text = para.ParagraphText;
                                //处理段落      
                                sb.Append(text + ",");
                            }
                        }
                    }
                }
                MessageBox.Show(sb.ToString());

            }
                
        }
    }
}
