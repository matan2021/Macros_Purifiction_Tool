using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;

namespace Macros_Purifiction_Tool
{
    public partial class Form1 : Form
    {
        private string file_path;
        private string desk_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            file_path = textBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = new Document();
                file_path = file_path.Replace("\"", "");
                doc.LoadFromFile(file_path);
                linkLabel1.Text = "";
                if (doc.IsContainMacro)
                {
                    doc.ClearMacros();
                }
                doc.SaveToFile(desk_path + "/Purged_Doc.docx", FileFormat.Docx);
                linkLabel1.Text = "Purged Succeeded, check your desktop";
            }
            catch (Exception) 
            {
                linkLabel1.Text = "Path does not exist! try again";
            }
        }
    }
        
}
