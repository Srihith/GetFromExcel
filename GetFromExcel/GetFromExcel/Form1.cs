using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace GetFromExcel
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        DataSet result;

        //Open Data
        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook| *.xls", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fs);
                    reader.IsFirstRowAsColumnNames = true;
                  
                    result = reader.AsDataSet();
                    cboSheet.Items.Clear();
                
                    foreach (DataTable dt in result.Tables)
                        {
                            cboSheet.Items.Add(dt.TableName);
                        }
                    
               

                    reader.Close();
                }
            }
        }

        //Updates Grid View
        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView.DataSource = result.Tables[cboSheet.SelectedIndex];
        
        }

        //Store Data
        int y;
        
        private void button1_Click(object sender, EventArgs e)
        {
 
            DataTable dt = dataGridView.DataSource as DataTable;


            Dictionary<string, int> dict = new Dictionary<string, int>();
            double[] count = new double[dt.Rows.Count];

            for (int x = 0; x <= dt.Rows.Count - 1; x++)
            {
                var src_user = dt.Rows[x]["Source User"].ToString();
                var hit_count = int.Parse(dt.Rows[x]["Count"].ToString());

                if (dict.ContainsKey(src_user))
                {
                    dict[src_user] = int.Parse(dict[src_user].ToString()) + hit_count;
                }
                else
                {
                    dict.Add(src_user, hit_count);
                }
            }

            var sortedDict = dict.OrderBy(x => x.Value);
            dt.AcceptChanges(); 
            StreamWriter File = new StreamWriter("Top100Users.txt");

            int users = 0;
            foreach (KeyValuePair<string, int> kvp in dict)
            {
                File.Write("{0},{1}\r\n", kvp.Key, kvp.Value);
                if (users > 99)
                {
                    break;
                }
                else
                {
                    users++;
                }
            }
            File.Close();





        }
        DataSet resultsAdd;
        //Add data
        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook| *.xls", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FileStream fsf = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fsf);
                    reader.IsFirstRowAsColumnNames = true;

                    resultsAdd = reader.AsDataSet();

                    comboBox.Items.Clear();
                    foreach (DataTable dtAdd in resultsAdd.Tables)
                    {
                        comboBox.Items.Add(dtAdd.TableName);
                    }



                    reader.Close();
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = resultsAdd.Tables[comboBox.SelectedIndex];

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            result.Tables[cboSheet.SelectedIndex].Merge(resultsAdd.Tables[comboBox.SelectedIndex]);
            result.Tables[0].DefaultView.Sort = "Count DESC";
            dataGridView.DataSource = result.Tables[cboSheet.SelectedIndex];
        }

       
    }
}
