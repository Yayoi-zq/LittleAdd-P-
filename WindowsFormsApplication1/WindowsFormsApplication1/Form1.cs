using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;//word库引用

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = System.IO.Path.GetFullPath(openFileDialog1.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = saveFileDialog1.FileName.ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            String docFileName = textBox1.Text;         //打开文件目录
            String resultFileName = textBox2.Text;      //输出文件目录
            try
            {
                // C#读取word文件之实例化
                object fileObj = docFileName;
                object nullObj = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.ApplicationClass wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref fileObj, ref nullObj, ref nullObj,
                    ref nullObj, ref nullObj, ref nullObj, ref nullObj, ref nullObj, ref nullObj, ref nullObj, ref nullObj,
                    ref nullObj, ref nullObj, ref nullObj, ref nullObj, ref nullObj);

                // 获取doc文件中的文本,并添加标签
                int count = doc.Paragraphs.Count;
                List<String> resultStr = new List<string>();

                //写入日志
                textBox3.Text = textBox3.Text + Environment.NewLine + "打开文件" + docFileName;

                for (int i = 1; i <= count; i++)
                {
                    string temp = doc.Paragraphs[i].Range.Text.Trim();
                    //写入日志
                    textBox3.Text = textBox3.Text + Environment.NewLine + "读取第" + i + "段：'" + temp + "'";
                    //类型判断
                    if (temp.Equals(""))
                    {
                        temp = "<br/>";
                    }
                    else
                    {
                        //添加p标签
                        temp = "<p>" + temp + "</p>";
                    }
                    resultStr.Add(temp);
                }

                //写入日志
                textBox3.Text = textBox3.Text + Environment.NewLine + "读取完毕";

                // 输出到文件result.txt
                FileStream fs = new FileStream(resultFileName, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                for (int i = 0; i < count; i++)
                {
                    sw.WriteLine(resultStr[i]);
                    sw.Flush();
                    //写入日志
                    textBox3.Text = textBox3.Text + Environment.NewLine + "正在写入第"+ (i+1) +"段";
                }
                sw.Close();
                fs.Close();

                //写入日志
                textBox3.Text = textBox3.Text + Environment.NewLine + "写入完毕，输出到" + resultFileName;
                textBox3.Text = textBox3.Text + Environment.NewLine + "========================";

                // 关闭文件
                doc.Close(ref nullObj, ref nullObj, ref nullObj);
                // 关闭COM
                wordApp.Quit(ref nullObj, ref nullObj, ref nullObj);

                return;
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.ToString());
            }
        }
    }
}
