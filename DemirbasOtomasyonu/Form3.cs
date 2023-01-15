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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti1 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Desktop\zimmet.xlsx;
Extended Properties='Excel 12.0 Xml; HDR=YES;'");
        void Veriler()
        {
            OleDbDataAdapter dataAdapter1 = new OleDbDataAdapter("Select * From [Sayfa1$]", baglanti1);
            DataTable dataTable1 = new DataTable();
            dataAdapter1.Fill(dataTable1);
            dataGridView2.DataSource = dataTable1;
        }
        

        private void Form3_Load(object sender, EventArgs e)
        {
            Veriler();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            baglanti1.Open();
            // sorgu yazıyoruz values bu değerleri nereden alacak
            OleDbCommand komut = new OleDbCommand("insert into [Sayfa1$] (PersonelİsimSoyisim,DemirbaşNo,TeslimTarihi,İadeTarihi,Durum) values (@p1,@p2,@p3,@p4,@p5)", baglanti1);
            // parametreleri atama işlemi
            komut.Parameters.AddWithValue("@p1", textBox1.Text);
            komut.Parameters.AddWithValue("@p2", textBox2.Text);
            komut.Parameters.AddWithValue("@p3", dateTimePicker1.Text);
            komut.Parameters.AddWithValue("@p4", dateTimePicker2.Text);
            komut.Parameters.AddWithValue("@p5", comboBox1.Text);
            // sorgumuzu çalıştırır
            komut.ExecuteNonQuery();
            baglanti1.Close();
            MessageBox.Show("Demirbaş Kaydedildi.");
            Veriler();
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            baglanti1.Open();
            OleDbCommand komut = new OleDbCommand("update [Sayfa1$] set  DemirbaşNo=@p2,TeslimTarihi=@p3,İadeTarihi=@p4,Durum=@p5 where PersonelİsimSoyisim=@p1", baglanti1);
            komut.Parameters.AddWithValue("@p2", textBox2.Text);
            komut.Parameters.AddWithValue("@p3", dateTimePicker1.Text);
            komut.Parameters.AddWithValue("@p4", dateTimePicker2.Text);
            komut.Parameters.AddWithValue("@p5", comboBox1.Text);
            komut.Parameters.AddWithValue("@p1", textBox1.Text);
            komut.ExecuteNonQuery();
            MessageBox.Show("Demirbaş güncellendi.");
            baglanti1.Close();
            Veriler();
        }

        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int secilen = dataGridView2.SelectedCells[0].RowIndex;
            //data gridin seçmiş olduğu satırın hücrelerinin içinde 0. hücresinin değerini string yap
            textBox1.Text = dataGridView2.Rows[secilen].Cells[0].Value.ToString();
            textBox2.Text = dataGridView2.Rows[secilen].Cells[1].Value.ToString();
            dateTimePicker1.Text = dataGridView2.Rows[secilen].Cells[2].Value.ToString();
            dateTimePicker2.Text = dataGridView2.Rows[secilen].Cells[3].Value.ToString();
            comboBox1.Text = dataGridView2.Rows[secilen].Cells[4].Value.ToString();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Visible = false;
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            
            
       }

        private void Button4_Click_1(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.Show();
            this.Visible = false;

        }
    }
}
