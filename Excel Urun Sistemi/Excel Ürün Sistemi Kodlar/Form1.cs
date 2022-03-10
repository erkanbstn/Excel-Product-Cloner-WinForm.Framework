using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using DevExpress.XtraEditors;
using Microsoft.Office.Interop.Excel;

namespace Excel_Ürün_Sistemi_Kodlar
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string nereye = "", neyinismi = "", bunu = "", excel = "";
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (Fbd.ShowDialog() == DialogResult.OK)
            {

                bunu = Fbd.SelectedPath.ToString();
                textEdit1.Text = bunu;
            }
            else
            {
                XtraMessageBox.Show("Lütfen Kopyalanacak Dosya Yolunu Belirleyiniz...", "Uyarı Sistemi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (Fbd2.ShowDialog() == DialogResult.OK)
            {
                nereye = Fbd2.SelectedPath.ToString();
                textEdit2.Text = nereye;
            }
            else
            {
                XtraMessageBox.Show("Dosyanın Kopyalanacağı Klasörü Seçmediniz...", "Uyarı Sistemi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        int kopyasayi;

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            XtraMessageBox.Show("0536 321 72 19 - Ulaşım Sağlayabilirsiniz...", "Bilgilendirme Mesajı", MessageBoxButtons.OK, MessageBoxIcon.Question);

        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            if (!textEdit1.Text.Equals("") && !textEdit2.Text.Equals("") && !textEdit3.Text.Equals(""))
            {
                neyinismi = Convert.ToString(dataGridView1.Columns[0].HeaderText.ToString());

                kopyasayi = Convert.ToInt32(dataGridView1.Columns[1].HeaderText.ToString());
                if (File.Exists(bunu + "\\" + neyinismi + ".jpg") == true)
                {
                    if (File.Exists(nereye + "\\" + neyinismi + ".jpg") == false)
                    {
                        File.Copy(bunu + "\\" + neyinismi + ".jpg", nereye + "\\" + neyinismi + ".jpg");

                    }

                }
                else
                {
                    listBox1.Items.Add(neyinismi + ".jpg");

                }


                for (int j = 0; j < dataGridView1.RowCount - 1; j++)
                {
                    if (j == 54)
                    {
                        kopyasayi = Convert.ToInt32(dataGridView1.Rows[j].Cells[1].Value.ToString());

                    }
                    else
                    {
                        kopyasayi = Convert.ToInt32(dataGridView1.Rows[j].Cells[1].Value.ToString());

                    }

                    for (int i = 1; i <= kopyasayi; i++)
                    {
                        neyinismi = dataGridView1.Rows[j].Cells[0].Value.ToString();
                        if (File.Exists(bunu + "\\" + neyinismi + ".jpg") == true)
                        {
                            if (File.Exists(nereye + "\\" + neyinismi + "(" + i.ToString() + ")" + ".jpg") == false)
                            {
                                File.Copy(bunu + "\\" + neyinismi + ".jpg", nereye + "\\" + neyinismi + "(" + i.ToString() + ")" + ".jpg");

                            }

                        }
                        else
                        {
                            listBox1.Items.Add(neyinismi + ".jpg");
                            goto geçla;
                        }



                    }
                geçla:;
                }
                XtraMessageBox.Show("Dosya Kopyalama İşlemi Başarılı", "Dosya Kopyalandı...", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            else
            {
                XtraMessageBox.Show("Lütfen Dosya Yolları Kısımlarını Doldurunuz...", "Uyarı Sistemi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            Opf1.Filter = "Xlsx Dosyaları |*.xlsx| Bütün Dosyalar|*.*";
            if (Opf1.ShowDialog() == DialogResult.OK)
            {
                excel = Opf1.FileName;
                textEdit3.Text = excel;
                simpleButton3_Click(sender, e);
            }
            else
            {
                XtraMessageBox.Show("Lütfen Excel Dosya Yolunu Belirleyiniz...", "Uyarı Sistemi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (Opf2.ShowDialog() == DialogResult.OK)
            {
                if (listBox1.Items.Count > 0)
                {


                    int adet = listBox1.Items.Count;
                    if (listBox1.Items.Count > 0)
                    {
                        Random rnd = new Random();
                        int sayi = rnd.Next();
                        string dosyayolu;
                        dosyayolu = Opf2.FileName /*+ "\\" + "Olmayan Resimler" + sayi.ToString() + ".xlsx"*/;

                        Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
                        excelapp.Visible = false;
                        object missing = Type.Missing;

                        Workbook calismakitabı = excelapp.Workbooks.Add(missing);
                        calismakitabı = excelapp.Workbooks.Open(dosyayolu);

                        Worksheet shet1 = (Worksheet)calismakitabı.Sheets[1];

                        int satir = 1;
                        int sütun = 1;
                        for (int i = 0; i < adet; i++)
                        {
                            for (int j = 0; j < 1; j++)
                            {
                                Range myrange1 = (Range)shet1.Cells[satir + i, sütun + j];
                                myrange1.Value2 = listBox1.Items[i] == null ? "" : listBox1.Items[i];
                                myrange1.Select();
                            }

                        }


                        object False = false;
                        object True = true;
                        XlFileFormat format = XlFileFormat.xlWorkbookDefault;
                        calismakitabı.SaveAs(dosyayolu, format, missing, missing, missing, False, XlSaveAsAccessMode.xlNoChange, missing, false, missing, missing, missing);
                        PrintDialog printdialog = new PrintDialog();
                        printdialog.PrintToFile = true;
                        calismakitabı.PrintOutEx(missing, missing, printdialog.PrinterSettings.Copies, false, printdialog.PrinterSettings.PrinterName, printdialog.PrinterSettings.PrintToFile, printdialog.PrinterSettings.Collate, printdialog.PrinterSettings.PrintFileName);
                        excelapp.Quit();
                    }
                }
                else
                {
                    XtraMessageBox.Show("Lütfen Aktarılacak Excel Dosyasını Seçiniz...", "Uyarı Sistemi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                XtraMessageBox.Show("Lütfen Aktarılacak Dosya İsimlerini Yazdırınız...", "Uyarı Sistemi", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
    }
}
