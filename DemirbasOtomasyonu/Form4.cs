using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace DemirbasOtomasyonu
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Desktop\zimmet.xlsx;
Extended Properties='Excel 12.0 Xml;'");
        void Arama()
        {
            OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;

        }
        private void Form4_Load(object sender, EventArgs e)
        {
            Arama();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            string aranan = textBox1.Text.Trim().ToUpper();
            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                foreach(DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach(DataGridViewCell cell in dataGridView1.Rows[i].Cells)
                    {
                        if(cell.Value!=null)
                        {
                            if(cell.Value.ToString().ToString()==aranan)
                            {
                                cell.Style.BackColor = Color.DarkTurquoise;
                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}
