using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Sempozyum_Günü_Kontrolü
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable dt = new DataTable();
        DataTable table = new DataTable();
        void griddoldur_kayitli()
        {
            OleDbConnection con;
            OleDbDataAdapter da;
            DataSet ds;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=veri.accdb");
            da = new OleDbDataAdapter("SElect *from kayitli", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "kayitli");
            metroGrid1_deneme.DataSource = ds.Tables["kayitli"];
            con.Close();
        }
        void griddoldur_kayitsiz()
        {
            OleDbConnection con;
            OleDbDataAdapter da;
            DataSet ds;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=veri.accdb");
            da = new OleDbDataAdapter("SElect *from kayitsiz", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "kayitsiz");
            metroGrid3_deneme.DataSource = ds.Tables["kayitsiz"];
            con.Close();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            table.Columns.Add("isim", typeof(string));

            griddoldur_kayitli();
            griddoldur_kayitsiz();

            if(metroGrid1_deneme.Rows.Count != 0)
            {
                OleDbConnection con;
                OleDbDataAdapter da;
                DataSet ds;
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=veri.accdb");
                da = new OleDbDataAdapter("SElect *from kayitli", con);
                ds = new DataSet();
                con.Open();
                da.Fill(ds, "kayitli");
                metroGrid1.DataSource = ds.Tables["kayitli"];
                con.Close();

                metroGrid1.Sort(this.metroGrid1.Columns[0], ListSortDirection.Ascending);

                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    if(metroGrid1.Rows[i].Cells[2].Value.ToString() == "+")
                    {
                        DataGridViewCellStyle style = new DataGridViewCellStyle();
                        style.BackColor = Color.DarkGreen;
                        style.ForeColor = Color.White;

                        metroGrid1.Rows[i].DefaultCellStyle = style;
                    }
                }

                metroGrid1.Columns[0].Width = 160;
                metroGrid1.Columns[0].HeaderText = "İsim - Soyisim";
                metroGrid1.Columns[1].Width = 190;
                metroGrid1.Columns[1].HeaderText = "Mail Adresleri";
                metroGrid1.Columns[2].Visible = false;

                metroGrid1.ClearSelection();

                metroLabel5.Text = metroGrid1.Rows.Count.ToString();

                metroGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }
            if(metroGrid3_deneme.Rows.Count != 0)
            {
                OleDbConnection con;
                OleDbDataAdapter da;
                DataSet ds;
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=veri.accdb");
                da = new OleDbDataAdapter("SElect *from kayitsiz", con);
                ds = new DataSet();
                con.Open();
                da.Fill(ds, "kayitsiz");
                metroGrid3.DataSource = ds.Tables["kayitsiz"];
                con.Close();

                metroGrid3.Columns[0].Width = 215;
                metroGrid3.Columns[0].HeaderText = "İsim - Soyisim";
                metroGrid3.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                metroGrid3.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                metroGrid3.ClearSelection();

                metroGrid3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }
        }
        private void btn_al_Click(object sender, EventArgs e)
        {
            metroGrid1.ClearSelection();

            OpenFileDialog openfile1 = new OpenFileDialog
            {
                Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                Title = "Veri Excel'ini seçiniz..."
            };
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = openfile1.FileName;
            }

            Excel.Application oXL = new Excel.Application(); //hmm demek nuget paketten bulmak gerekiyormuş seni ve sonrada öyle using Excel diyerek kullanmak gerekiyormuş
            if (textBox1.Text == string.Empty)
            {
                return;
            }
            else
            {
                if(metroGrid1.Rows.Count != 0)
                {
                    dt.Rows.Clear();
                }

                Excel.Workbook oWB = oXL.Workbooks.Open(textBox1.Text); // hata burada oluşuyor demek

                List<string> liste = new List<string>();
                foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                {
                    liste.Add(oSheet.Name);
                }
                oWB.Close();
                oXL.Quit();
                oWB = null;
                oXL = null;
                metroGrid2.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();
                textBox2.Text = metroGrid2.Rows[0].Cells[0].Value.ToString();

                OleDbCommand komut = new OleDbCommand();
                string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                OleDbConnection conn = new OleDbConnection(pathconn);
                OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
                MyDataAdapter.Fill(dt);
                metroGrid1.DataSource = dt;

                metroGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;

                metroGrid1.Columns[0].Width = 160;
                metroGrid1.Columns[0].HeaderText = "İsim - Soyisim";
                metroGrid1.Columns[1].Width = 190;
                metroGrid1.Columns[1].HeaderText = "Mail Adresleri";
                metroGrid1.Columns[2].Visible = false;

                metroGrid1.ClearSelection();

                metroLabel5.Text = metroGrid1.Rows.Count.ToString();
                //try catch atarak 0 eleman varken tıklamayı çözelim, gerçi onla da çözülmüyordu sanırım ya, tek yol onu gizlemek :/
                //değilmiş meğerse, cell clik yapınca olay çözüldü

                OleDbCommand komut2;
                OleDbCommand komut3;
                OleDbCommand komut4;
                string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
                OleDbConnection baglanti = new OleDbConnection(vtyolu);

                for (int i = 0; i < metroGrid1_deneme.Rows.Count; i++)
                {
                    string h = metroGrid1_deneme.Rows[i].Cells[1].Value.ToString();

                    baglanti.Open();
                    string sil = "delete from kayitli where mail=@mail";
                    komut3 = new OleDbCommand(sil, baglanti);
                    komut3.Parameters.AddWithValue("@mail", h);
                    komut3.ExecuteNonQuery();
                    komut3.Dispose();
                    baglanti.Close();
                }

                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    baglanti.Open();
                    string ekle = "insert into kayitli(isim,mail,arti) values (@isim,@mail,@arti)";
                    komut2 = new OleDbCommand(ekle, baglanti);
                    komut2.Parameters.AddWithValue("@isim", metroGrid1.Rows[i].Cells[0].Value.ToString());
                    komut2.Parameters.AddWithValue("@mail", metroGrid1.Rows[i].Cells[1].Value.ToString());
                    komut2.Parameters.AddWithValue("@arti", metroGrid1.Rows[i].Cells[2].Value.ToString());
                    komut2.ExecuteNonQuery();
                    komut2.Dispose();
                    baglanti.Close();
                }

                for (int i = 0; i < metroGrid3.Rows.Count; i++)
                {
                    string h = metroGrid3.Rows[i].Cells[0].Value.ToString();

                    baglanti.Open();
                    string sil = "delete from kayitsiz where isim=@isim";
                    komut4 = new OleDbCommand(sil, baglanti);
                    komut4.Parameters.AddWithValue("@isim", h);
                    komut4.ExecuteNonQuery();
                    komut4.Dispose();
                    baglanti.Close();
                }

                OleDbConnection con;
                OleDbDataAdapter da;
                DataSet ds;
                con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=veri.accdb");
                da = new OleDbDataAdapter("SElect *from kayitsiz", con);
                ds = new DataSet();
                con.Open();
                da.Fill(ds, "kayitsiz");
                metroGrid3.DataSource = ds.Tables["kayitsiz"];
                con.Close();

                metroGrid3.Columns[0].Width = 215;
                metroGrid3.Columns[0].HeaderText = "İsim - Soyisim";
                metroGrid3.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                metroGrid3.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                metroGrid3.ClearSelection();

                metroGrid3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }
        }
        private void metroTextBox1_TextChanged(object sender, EventArgs e)
        {
            metroGrid1.ClearSelection();

            if(dt.Rows.Count == 0)
            {
                dt.Columns.Add("isim", typeof(string));
                dt.Columns.Add("mail", typeof(string));
                dt.Columns.Add("arti", typeof(string));

                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    dt.Rows.Add(metroGrid1.Rows[i].Cells[0].Value.ToString(), metroGrid1.Rows[i].Cells[1].Value.ToString(), metroGrid1.Rows[i].Cells[2].Value.ToString());
                }
            }

            DataView dv = dt.DefaultView;
            dv.RowFilter = "isim LIKE '" + metroTextBox1.Text + "%'";
            metroGrid1.DataSource = dv;

            for (int i = 0; i < metroGrid1.Rows.Count; i++)
            {
                if (metroGrid1.Rows[i].Cells[2].Value.ToString() == "+")
                {
                    DataGridViewCellStyle style = new DataGridViewCellStyle();
                    style.BackColor = Color.DarkGreen;
                    style.ForeColor = Color.White;

                    metroGrid1.Rows[i].DefaultCellStyle = style;
                }
            }

            metroGrid1.ClearSelection();
        }
        private void btn_ver_Click(object sender, EventArgs e)
        {
            if (metroTextBox1.Text == "")
            {
                MessageBox.Show("Lütfen kutucuğa bir isim giriniz", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (metroGrid1.Rows.Count != 0)
            {
                MessageBox.Show(metroTextBox1.Text.ToString() + " kelimesini içeren bazı kayıtların olduğu hala görünebiliyor. Lütfen kişinin adını tam olarak yazınız ve butona öyle basınız. Eğer kişi kayıtlıysa lütfen sadece üstüne tıklayın ve diğer isimlerin kontrolüne geçiniz","Hata",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
            else
            {
                metroGrid3.ClearSelection();
                metroGrid1.ClearSelection();

                if (table.Rows.Count == 0 && metroGrid3.Rows.Count != 0)
                {
                    for (int i = 0; i < metroGrid3.Rows.Count; i++)
                    {
                        table.Rows.Add(metroGrid3.Rows[i].Cells[0].Value.ToString());
                    }
                }

                if (metroGrid1.Rows.Count == 0)
                {
                    table.Rows.Add(metroTextBox1.Text.ToString());
                }
                metroGrid3.DataSource = table;

                metroGrid1.ClearSelection();
                metroGrid3.ClearSelection();

                metroGrid3.Columns[0].Width = 215;
                metroGrid3.Columns[0].HeaderText = "İsim - Soyisim";
                metroGrid3.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                metroGrid3.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                metroGrid3.BorderStyle = System.Windows.Forms.BorderStyle.None;
                //kırmızı beyaz olanlar katılımcılar, diğer taraftakiler aslında para yatırmayanlardan oluşuyor

                OleDbCommand komut3;
                string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
                OleDbConnection baglanti = new OleDbConnection(vtyolu);

                baglanti.Open();
                string ekle = "insert into kayitsiz(isim) values (@isim)";
                komut3 = new OleDbCommand(ekle, baglanti);
                komut3.Parameters.AddWithValue("@isim", metroTextBox1.Text.ToString());
                komut3.ExecuteNonQuery();
                komut3.Dispose();
                baglanti.Close();
            }
        }
        private void metroGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //vay aq, cell click deyince demek illa ki hücre olma şartını arıyormuş, bunu denediğim ve öğrendiğim iyi oldu
            OleDbCommand komut;
            OleDbCommand komut2;
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);

            metroGrid1.ClearSelection();

            if (metroCheckBox1.Checked == true)
            {
                DataGridViewCellStyle style = new DataGridViewCellStyle();
                style.BackColor = default;
                style.ForeColor = default;

                int x = metroGrid1.CurrentRow.Index;

                metroGrid1.Rows[x].DefaultCellStyle = style;
                metroGrid1.Rows[x].Cells[2].Value = "";

                baglanti.Open();
                string guncelle = "update kayitli set arti=@arti where mail=@mail";
                komut = new OleDbCommand(guncelle, baglanti);
                komut.Parameters.AddWithValue("@arti", metroGrid1.Rows[x].Cells[2].Value.ToString());
                komut.Parameters.AddWithValue("@mail", metroGrid1.Rows[x].Cells[1].Value.ToString());
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
            }
            else
            {
                DataGridViewCellStyle style = new DataGridViewCellStyle();
                style.BackColor = Color.DarkGreen;
                style.ForeColor = Color.White;

                int x = metroGrid1.CurrentRow.Index;

                metroGrid1.Rows[x].DefaultCellStyle = style;
                metroGrid1.Rows[x].Cells[2].Value = "+";

                baglanti.Open();
                string guncelle = "update kayitli set arti=@arti where mail=@mail";
                komut2 = new OleDbCommand(guncelle, baglanti);
                komut2.Parameters.AddWithValue("@arti", metroGrid1.Rows[x].Cells[2].Value.ToString());
                komut2.Parameters.AddWithValue("@mail", metroGrid1.Rows[x].Cells[1].Value.ToString());
                komut2.ExecuteNonQuery();
                komut2.Dispose();
                baglanti.Close();
            }

            metroGrid1.ClearSelection();
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            OleDbCommand komut;
            OleDbCommand komut2;
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);

            for (int i = 0; i < metroGrid1.Rows.Count; i++)
            {
                if (metroGrid1.Rows[i].Cells[2].Value.ToString() == "+")
                {
                    DataGridViewCellStyle style = new DataGridViewCellStyle();
                    style.BackColor = default;
                    style.ForeColor = default;

                    metroGrid1.Rows[i].DefaultCellStyle = style;
                    metroGrid1.Rows[i].Cells[2].Value = "";
                }
            }

            for (int i = 0; i < metroGrid1.Rows.Count; i++)
            {
                baglanti.Open();
                string guncelle = "update kayitli set arti=@arti where mail=@mail";
                komut = new OleDbCommand(guncelle, baglanti);
                komut.Parameters.AddWithValue("@arti", metroGrid1.Rows[i].Cells[2].Value.ToString());
                komut.Parameters.AddWithValue("@mail", metroGrid1.Rows[i].Cells[1].Value.ToString());
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
            }

            for (int i = 0; i < metroGrid3.Rows.Count; i++)
            {
                string h = metroGrid3.Rows[i].Cells[0].Value.ToString();

                baglanti.Open();
                string sil = "delete from kayitsiz where isim=@isim";
                komut2 = new OleDbCommand(sil, baglanti);
                komut2.Parameters.AddWithValue("@isim", h);
                komut2.ExecuteNonQuery();
                komut2.Dispose();
                baglanti.Close();
            }

            table.Rows.Clear();

            OleDbConnection con;
            OleDbDataAdapter da;
            DataSet ds;
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=veri.accdb");
            da = new OleDbDataAdapter("SElect *from kayitsiz", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "kayitsiz");
            metroGrid3.DataSource = ds.Tables["kayitsiz"];
            con.Close();

            metroGrid3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;

            //hee, bak metro1'i silmeyince o artılar hafızada kaldı tabi haliyle, demek silmek lazım onu o zaman
            //aa yok lan, silmek değilde + ları silip geri grid doldurtmak lazım
            //heee doğru ya, silmek değil upgrade yapmam lazım abi benim unuttum birden ya

            metroGrid3.Columns[0].Width = 215;
            metroGrid3.Columns[0].HeaderText = "İsim - Soyisim";
            metroGrid3.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid3.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            metroGrid3.ClearSelection();

            metroGrid3.BorderStyle = System.Windows.Forms.BorderStyle.None;

            MessageBox.Show("Program yeni oturum için hazır durumdadır", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroGrid1_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }
    }
}
