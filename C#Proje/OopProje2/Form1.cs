using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace OopProje2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int sutun = dataGridView1.Columns.Count;
            int satir = dataGridView1.Rows.Count;
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < sutun; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < satir; i++)
            {
                for (int j = 0; j < sutun; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();


                }
            }
        }
        string koltukno = "";
        public string KoltukSec()
        {
            if (checkBox1.Checked)
                koltukno = "1";
            if (checkBox2.Checked)
                koltukno = "2";
            if (checkBox3.Checked)
                koltukno = "3";
            if (checkBox4.Checked)
                koltukno = "4";
            if (checkBox5.Checked)
                koltukno = "5";
            if (checkBox6.Checked)
                koltukno = "6";
            if (checkBox7.Checked)
                koltukno = "7";
            if (checkBox8.Checked)
                koltukno = "8";
            if (checkBox9.Checked)
                koltukno = "9";
            if (checkBox10.Checked)
                koltukno = "10";
            if (checkBox11.Checked)
                koltukno = "11";
            if (checkBox12.Checked)
                koltukno = "12";
            if (checkBox13.Checked)
                koltukno = "13";
            if (checkBox14.Checked)
                koltukno = "14";
            if (checkBox15.Checked)
                koltukno = "15";
            if (checkBox16.Checked)
                koltukno = "16";
            if (checkBox17.Checked)
                koltukno = "17";
            if (checkBox18.Checked)
                koltukno = "18";
            return koltukno;
        }
        void GridveTCAyarla()
        {
            textBox1.MaxLength = 11;
            dataGridView1.ColumnCount = 10;
            dataGridView1.Columns[0].Name = "TC";
            dataGridView1.Columns[1].Name = "Ad Soyad";
            dataGridView1.Columns[2].Name = "Telefon";
            dataGridView1.Columns[3].Name = "Mail";
            dataGridView1.Columns[4].Name = "Kalkış Şehri";
            dataGridView1.Columns[5].Name = "Varış Şehri";
            dataGridView1.Columns[6].Name = "Kalkış Saati";
            dataGridView1.Columns[7].Name = "Varış Saati";
            dataGridView1.Columns[8].Name = "KoltukNo";
            dataGridView1.Columns[9].Name = "Bilet Fiyat";
        }
        public static bool TcDogrula(string tcKimlikNo)
        {
            bool returnvalue = false;
            if (tcKimlikNo.Length == 11)
            {
                Int64 ATCNO, BTCNO, TcNo;
                long C1, C2, C3, C4, C5, C6, C7, C8, C9, Q1, Q2;
                TcNo = Int64.Parse(tcKimlikNo);
                ATCNO = TcNo / 100;
                BTCNO = TcNo / 100;
                C1 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C2 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C3 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C4 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C5 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C6 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C7 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C8 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                C9 = ATCNO % 10;
                ATCNO = ATCNO / 10;
                Q1 = ((10 - ((((C1 + C3 + C5 + C7 + C9) * 3) + (C2 + C4 + C6 + C8)) % 10)) % 10);
                Q2 = ((10 - (((((C2 + C4 + C6 + C8) + Q1) * 3) + (C1 + C3 + C5 + C7 + C9)) % 10)) % 10);
                returnvalue = ((BTCNO * 100) + (Q1 * 10) + Q2 == TcNo);
            }

            return returnvalue;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] Şehirler = {"Adana","Adıyaman", "Afyon", "Ağrı", "Amasya", "Ankara", "Antalya", "Artvin",
            "Aydın", "Balıkesir","Bilecik", "Bingöl", "Bitlis", "Bolu", "Burdur", "Bursa", "Çanakkale",
            "Çankırı", "Çorum","Denizli","Diyarbakır", "Edirne", "Elazığ", "Erzincan", "Erzurum ", "Eskişehir",
            "Gaziantep", "Giresun","Gümüşhane", "Hakkari", "Hatay", "Isparta", "Mersin", "İstanbul", "İzmir",
            "Kars", "Kastamonu", "Kayseri","Kırklareli", "Kırşehir", "Kocaeli", "Konya", "Kütahya ", "Malatya",
            "Manisa", "Kahramanmaraş", "Mardin", "Muğla", "Muş", "Nevşehir", "Niğde", "Ordu", "Rize", "Sakarya",
            "Samsun", "Siirt", "Sinop", "Sivas", "Tekirdağ", "Tokat", "Trabzon", "Tunceli", "Şanlıurfa", "Uşak",
            "Van", "Yozgat", "Zonguldak", "Aksaray", "Bayburt", "Karaman", "Kırıkkale", "Batman", "Şırnak",
            "Bartın", "Ardahan", "Iğdır", "Yalova", "Karabük ", "Kilis", "Osmaniye ", "Düzce"};
            comboBox1.Items.AddRange(Şehirler);
            comboBox2.Items.AddRange(Şehirler);

            string[] Saatler = { "00:00", "00:30", "01:00", "01:30", "02:00", "02:30", "03:00", "03:30", "04:00",
            "04:30","05:00","05:30","06:00","06:30","07:00","07:30","08:00","08:30","09:00","09:30","10:00","10:30",
            "11:00","11:30","12:00","12:30","13:00","13:30","14:00","14:30","15:00","15:30","16:00","16:30","17:00",
            "17:30","18:00","18:30","19:00","19:30","20:00","20:30","21:00","21:30","22:00","22:30","23:00","23:30" };
            comboBox3.Items.AddRange(Saatler);
            comboBox4.Items.AddRange(Saatler);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string tc, adsoyad, tel, mail, kalkissehir, varissehir, kalkissaat, varissaat, koltukno;
            int bilet;
            tc = textBox1.Text;
            if (TcDogrula(tc) == true)
            {
                adsoyad = textBox2.Text;
                tel = textBox3.Text;
                mail = textBox4.Text;
                koltukno = textBox5.Text;
                bilet = Convert.ToInt32(textBox6.Text);
                kalkissehir = comboBox1.Text;
                varissehir = comboBox2.Text;
                kalkissaat = comboBox3.Text;
                varissaat = comboBox4.Text;
                dataGridView1.Rows.Add(tc, adsoyad, tel, mail, kalkissehir, varissehir, kalkissaat, varissaat, koltukno, bilet);
                MessageBox.Show("Yolcu Bileti Oluşturuldu İyi Uçuşlar!", "Sistem Bilgilendirme Mesajı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Lütfen Geçerli Bir Kimlik Numarası Giriniz.", "Sistem Uyarı Mesajı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Koltuğunuz Başarıyla Seçilmiştir.", "Bilgi");
            textBox5.Text = KoltukSec();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www.turkishairlines.com/tr-tr/");
        }
    }
}
