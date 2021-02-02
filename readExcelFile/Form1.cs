using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace readExcelFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = String.Empty;
            string fileExt = string.Empty;

            OpenFileDialog fileDialog = new OpenFileDialog();

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = fileDialog.FileName;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable table = new System.Data.DataTable();

                        table = ReadExcel(filePath, fileExt);
                        dataGridView1.DataSource = table;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("" + ex);
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public System.Data.DataTable ReadExcel(string path, string ext)
        {
            string conn = String.Empty;

            System.Data.DataTable table = new System.Data.DataTable();
            if (ext.CompareTo(".xls") == 0)
            {
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007
            }
            else
            {
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007
            }
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    System.Data.DataTable tExcelsheetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    string str5 = "[" + tExcelsheetName.Rows[0]["TABLE_NAME"].ToString().Replace("'", "") + "]";
                    MessageBox.Show(str5);

                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from " + str5, con);
                    adapter.Fill(table);
                }
                catch (Exception e)
                {
                    MessageBox.Show("" + e);
                }
            }
            return table;
        }
    }
}
