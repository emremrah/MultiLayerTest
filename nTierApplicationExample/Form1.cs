using nTierApplicationExample.UserControls;
using nTierApplicationExample.ValuesLayer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BLL = nTierApplicationExample.BusinessLogicLayer.BusinessLogicLayer;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office.Interop;


namespace nTierApplicationExample
{
    public partial class Form1 : Form
    {
        public Sehir newSehir;
        public SehirControl sehirControl = new SehirControl();
        public Kisi newKisi;
        public KisiControl kisiControl = new KisiControl();
        public Ilce newIlce;
        public IlceControl ilceControl = new IlceControl();

        public Form1 ()
        {
            InitializeComponent();
        }

        //Select
        private void button1_Click (object sender, EventArgs e)
        {
            BLL bll = new BLL();
            dataGridView1.DataSource = bll.Get(comboBox1.SelectedItem.ToString());
        }

        //Insert
        private void button2_Click (object sender, EventArgs e)
        {
            try {
                BLL BLL = new BLL();    //BLL Business Logic Layer
                if (comboBox1.SelectedItem.ToString() == "Sehir") {
                    Sehir newSehir = new Sehir(Convert.ToInt16(sehirControl.id.Text), sehirControl.ad.Text);
                    BLL.Insert<Sehir>(newSehir);
                    dataGridView1.DataSource = BLL.Get("Sehir");
                } else if (comboBox1.SelectedItem.ToString() == "Kisi") {
                    Kisi newKisi = new Kisi(Convert.ToInt16(kisiControl.id.Text), kisiControl.ad.Text, kisiControl.soyad.Text, Convert.ToInt16(kisiControl.yas.Text), kisiControl.adres.Text, kisiControl.sehir.Text, kisiControl.ilce.Text);
                    BLL.Insert<Kisi>(newKisi);
                    dataGridView1.DataSource = BLL.Get("Kisi");
                } else if (comboBox1.SelectedItem.ToString() == "Ilce") {
                    Ilce newIlce = new Ilce(Convert.ToInt16(ilceControl.id.Text), ilceControl.ad.Text, Convert.ToInt16(ilceControl.sehirId.Text));
                    BLL.Insert<Ilce>(newIlce);
                    dataGridView1.DataSource = BLL.Get("Ilce");
                }
            } catch {
                MessageBox.Show("An error occured during insert");
            }
        }

        //Update
        private void button3_Click (object sender, EventArgs e)
        {
            try {
                if (comboBox1.SelectedIndex == 0) {
                    BLL bll = new BLL();
                    Sehir newSehir = new Sehir(Convert.ToInt16(sehirControl.id.Text), sehirControl.ad.Text);

                    bll.Update<Sehir>(newSehir);
                    dataGridView1.DataSource = bll.Get("Sehir");
                } else if (comboBox1.SelectedItem.ToString() == "Kisi") {
                    BLL bll = new BLL();
                    Kisi newKisi = new Kisi(Convert.ToInt16(kisiControl.id.Text), kisiControl.ad.Text, kisiControl.soyad.Text, Convert.ToInt16(kisiControl.yas.Text), kisiControl.adres.Text, kisiControl.sehir.Text, kisiControl.ilce.Text);

                    bll.Update<Kisi>(newKisi);
                    dataGridView1.DataSource = bll.Get("Kisi");
                } else if (comboBox1.SelectedItem.ToString() == "Ilce") {
                    BLL bll = new BLL();
                    Ilce newIlce = new Ilce(Convert.ToInt16(ilceControl.id.Text), ilceControl.ad.Text, Convert.ToInt16(ilceControl.sehirId.Text));
                    bll.Update<Ilce>(newIlce);
                    dataGridView1.DataSource = bll.Get("Ilce");
                }
            } catch {
                MessageBox.Show("Error occured during update.");
            }
        }

        //Delete
        private void button4_Click (object sender, EventArgs e)
        {
            int deleteId = 0;
            if (comboBox1.SelectedItem.ToString() == "Sehir") {
                deleteId = Convert.ToInt16(sehirControl.id);
            }
            int[] deleteIndexes;
            try
            {
                Excel.Application xlApplication = (Excel.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel.Workbook xlWorkbook = (Excel.Workbook) xlApplication.ActiveWorkbook;
                Excel.Worksheet xlWorksheet = (Excel.Worksheet) xlWorkbook.ActiveSheet;
                Excel.Range xlRange = xlWorksheet.UsedRange;

                for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                    if (xlWorksheet.Cells[i+1, 1] = 1) {
                        for (int k = 0; k < dataGridView1.Columns.Count; k++) {
                            xlWorksheet.Cells[i + 1, k] = null;
                        }
                    }
                    //for (int j = 0; j < dataGridView1.Columns.Count; j++) {
                    //    xlWorksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    //    xlWorksheet.Cells[3, 4] = null;
                    //}
                }
            }
            catch
            {
                MessageBox.Show("An error occured during deletion");   
            }
        }

        private void Form1_Load (object sender, EventArgs e)
        {
            comboBox1.Items.Add("Sehir");
            comboBox1.Items.Add("Kisi");
            comboBox1.Items.Add("Ilce");
            this.Controls.Add(sehirControl);
            this.Controls.Add(kisiControl);
            this.Controls.Add(ilceControl);
            foreach (Control c in this.Controls)
                if (c is UserControl) c.Visible = false;
        }

        //Combobox seçimi ve tablonun gösterilmesi
        private void comboBox1_SelectedValueChanged (object sender, EventArgs e)
        {
            foreach (Control u in this.Controls) {
                if (u.Name == comboBox1.SelectedItem + "Control") {
                    u.Visible = true;
                    u.Location = new Point(400, 0);
                    u.BringToFront();
                } else if (u is UserControl) u.Visible = false;
            }
            
            Get(comboBox1.SelectedItem.ToString());
            //if (comboBox1.SelectedIndex==0)
            //{
            //    this.Controls.Add(sehirControl);
            //    sehirControl.Location = new Point(400, 0);
            //    sehirControl.BringToFront();
            //    sehirControl.Visible = true;
            //    ilceControl.Visible = false;
            //    kisiControl.Visible = false;
            //    Get();
            //}
            //else if (comboBox1.SelectedIndex == 1)
            //{
            //    this.Controls.Add(kisiControl);
            //    kisiControl.Location = new Point(400, 0);
            //    kisiControl.BringToFront();
            //    sehirControl.Visible = false;
            //    ilceControl.Visible = false;
            //    kisiControl.Visible = true;
            //    Get();
            //}
            //else if (comboBox1.SelectedIndex == 2)
            //{
            //    this.Controls.Add(ilceControl);
            //    ilceControl.Location = new Point(400, 0);
            //    ilceControl.BringToFront();
            //    kisiControl.Visible = false;
            //    sehirControl.Visible = false;
            //    ilceControl.Visible = true;
            //    Get();
            //}
        }

        private void Get (string type)
        {
            try {
                BLL bll = new BLL();

                //BLL'deki GET fonksiyonu bir dataTable döndürüyor. DataTable da dataGridView için source
                //olarak kullanılıyor.
                //if (comboBox1.SelectedIndex == 0)
                //{
                //    dataGridView1.DataSource = bll.Get("Sehir");
                //}
                //else if (comboBox1.SelectedIndex == 1)
                //{
                //    dataGridView1.DataSource = bll.Get("Kisi");
                //}
                //else if (comboBox1.SelectedIndex == 2)
                //{
                //    dataGridView1.DataSource = bll.Get("Ilce");
                //}
                dataGridView1.DataSource = bll.Get(type);
            } catch {
                MessageBox.Show("An error occured.");
            }
        }

        //Excel'e datagridview'den veri göndermek için!
        private void button5_Click (object sender, EventArgs e)
        {
            Excel.Application xlApplication = (Excel.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook xlWorkbook = (Excel.Workbook) xlApplication.ActiveWorkbook;
            Excel.Worksheet xlWorksheet = (Excel.Worksheet) xlWorkbook.ActiveSheet;
            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                for (int j = 0; j < dataGridView1.Columns.Count; j++) {
                    xlWorksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    xlWorksheet.Cells[3, 4] = null;
                }
            }
            //dynamic allDataRange = xlWorksheet.UsedRange;
            //allDataRange.Sort(allDataRange.Columns[1], Excel.XlSortOrder.xlAscending);
        }

        private void dataGridView1_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e) {
            dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);

            Excel.Application xlApplication = (Excel.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook xlWorkbook = (Excel.Workbook) xlApplication.ActiveWorkbook;
            Excel.Worksheet xlWorksheet = (Excel.Worksheet) xlWorkbook.ActiveSheet;
            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                for (int j = 0; j < dataGridView1.Columns.Count; j++) {
                    xlWorksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
        }
    }
}
