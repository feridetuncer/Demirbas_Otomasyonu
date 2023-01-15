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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Desktop\demirbaş.xlsx;
Extended Properties='Excel 12.0 Xml; HDR=YES;'");

        void Veriler()
        {
            // sorgumu yazabilmnek için kullanılcak alanları seçiyor bağlantı nesnesinin bulunduğu adreste
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter("Select * From [Sayfa1$]", baglanti);
            //
            DataTable dataTable = new DataTable();
            // bu tablonun içini sorgudan gelen değerle doldur
            dataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
        }
            private void Form2_Load(object sender, EventArgs e)
        {
            Veriler();
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            baglanti.Open();
            // sorgu yazıyoruz values bu değerleri nereden alacak
            OleDbCommand komut = new OleDbCommand("insert into [Sayfa1$] (DemirbaşNo,SeriNo,MalzemeKod,MalzemeAdı,GirişTarihi,Durum) " +
                "values (@p1,@p2,@p3,@p4,@p5,@p6)", baglanti);
            // parametreleri atama işlemi
            komut.Parameters.AddWithValue("@p1", textBox1.Text);
            komut.Parameters.AddWithValue("@p2", textBox2.Text);
            komut.Parameters.AddWithValue("@p3", textBox3.Text);
            komut.Parameters.AddWithValue("@p4", textBox4.Text);
            komut.Parameters.AddWithValue("@p5", dateTimePicker1.Text);
            komut.Parameters.AddWithValue("@p6", comboBox1.Text);
            // sorgumuzu çalıştırır
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Demirbaş Kaydedildi.");
            Veriler();
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("update [Sayfa1$] set  SeriNo=@p2, MalzemeKod=@p3, MalzemeAdı=@p4, GirişTarihi=@p5," +
                " Durum=@p6 where DemirbaşNo=@p1", baglanti);
            komut.Parameters.AddWithValue("@p2", textBox2.Text);
            komut.Parameters.AddWithValue("@p3", textBox3.Text);
            komut.Parameters.AddWithValue("@p4", textBox4.Text);
            komut.Parameters.AddWithValue("@p5", dateTimePicker1.Text);
            komut.Parameters.AddWithValue("@p6", comboBox1.Text);
            komut.Parameters.AddWithValue("@p1", textBox1.Text);
            komut.ExecuteNonQuery();
            MessageBox.Show("Demirbaş güncellendi.");
            baglanti.Close();
            Veriler();
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //seçilen satırın indexini al
            int secilen = dataGridView1.SelectedCells[0].RowIndex;
            //data gridin seçmiş olduğu satırın hücrelerinin içinde 0. hücresinin değerini string yap
            textBox1.Text = dataGridView1.Rows[secilen].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.Rows[secilen].Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.Rows[secilen].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.Rows[secilen].Cells[3].Value.ToString();
            dateTimePicker1.Text = dataGridView1.Rows[secilen].Cells[4].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[secilen].Cells[5].Value.ToString();

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Visible = false;
        }

        private void Label7_Click(object sender, EventArgs e)
        {

        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
            this.Visible = false;

        }
    }
}
