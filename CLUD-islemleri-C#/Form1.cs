using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApplication9
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection bag = new OleDbConnection();
            bag.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + Application.StartupPath + "\\VT.accdb";
            bag.Open();
           
            string pno, ad, soyad, telefon,maas;
            pno = textBox1.Text;
            ad = textBox2.Text;
            soyad = textBox3.Text;
            telefon = textBox4.Text;
            maas = textBox5.Text;

            
            string sql = "insert into PERSONEL(PNO,AD,SOYAD,TELEFON,MAAS)";
            sql = sql + " values("+pno+",'"+ad+"','"+soyad+"','"+telefon+"',"+maas+")";
            OleDbCommand komut = new OleDbCommand(sql,bag);
            int sonuc = komut.ExecuteNonQuery();
            MessageBox.Show("Kayıt Eklendi");
            bag.Close();



        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            listBox5.Items.Clear();

            OleDbConnection bag = new OleDbConnection();
            bag.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + Application.StartupPath + "\\VT.accdb";
            bag.Open();

            string sql = "select * from PERSONEL";
            OleDbCommand komut = new OleDbCommand(sql, bag);
            OleDbDataReader oku = komut.ExecuteReader();
            while (oku.Read())
            {
                listBox1.Items.Add(oku["PNO"].ToString());
                listBox2.Items.Add(oku["AD"].ToString());
                listBox3.Items.Add(oku["SOYAD"].ToString());
                listBox4.Items.Add(oku["TELEFON"].ToString());
                listBox5.Items.Add(oku[4].ToString());
            }
            bag.Close();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int sira = listBox1.SelectedIndex;
            listBox2.SelectedIndex = sira;
            listBox3.SelectedIndex = sira;
            listBox4.SelectedIndex = sira;
            listBox5.SelectedIndex = sira;


            textBox1.Text = listBox1.Text;
            textBox2.Text = listBox2.Text;
            textBox3.Text = listBox3.Text;
            textBox4.Text = listBox4.Text;
            textBox5.Text = listBox5.Text;

        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            string silinecek = listBox1.Text;
            
            OleDbConnection bag = new OleDbConnection();
            bag.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + Application.StartupPath + "\\VT.accdb";
            bag.Open();

            string sql = "delete  from PERSONEL where PNO="+silinecek;
            OleDbCommand komut = new OleDbCommand(sql, bag);
            int sonuc = komut.ExecuteNonQuery();
            MessageBox.Show("kayıt silindi");
            bag.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OleDbConnection bag = new OleDbConnection();
            bag.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + Application.StartupPath + "\\VT.accdb";
            bag.Open();

            string sql = "update PERSONEL set AD=@T1, SOYAD=@T2,TELEFON=@T3, MAAS=@T4 where PNO=@T5";

            OleDbCommand komut = new OleDbCommand(sql, bag);
            komut.Parameters.Add("T1", textBox2.Text);
            komut.Parameters.Add("T2", textBox3.Text);
            komut.Parameters.Add("T3", textBox4.Text);
            komut.Parameters.Add("T4", textBox5.Text);
            komut.Parameters.Add("T5", textBox1.Text);
            int sonuc = komut.ExecuteNonQuery();
            MessageBox.Show("Kayıt Değiştirildi");
            bag.Close();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.Show();
        }
    }
}
