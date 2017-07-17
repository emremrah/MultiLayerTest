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
            dataGridView.DataSource = bll.Get(cbSelect.SelectedItem.ToString());
        }

        //Insert
        private void button2_Click (object sender, EventArgs e)
        {
            try {
                BLL BLL = new BLL();    //BLL Business Logic Layer
                if (cbSelect.SelectedItem.ToString() == "Sehir") {
                    Sehir newSehir = new Sehir(Convert.ToInt16(sehirControl.id.Text), sehirControl.ad.Text);
                    BLL.Insert<Sehir>(newSehir);
                    dataGridView.DataSource = BLL.Get("Sehir");
                } else if (cbSelect.SelectedItem.ToString() == "Kisi") {
                    Kisi newKisi = new Kisi(Convert.ToInt16(kisiControl.id.Text), kisiControl.ad.Text, kisiControl.soyad.Text, Convert.ToInt16(kisiControl.yas.Text), kisiControl.adres.Text, kisiControl.sehir.Text, kisiControl.ilce.Text);
                    BLL.Insert<Kisi>(newKisi);
                    dataGridView.DataSource = BLL.Get("Kisi");
                } else if (cbSelect.SelectedItem.ToString() == "Ilce") {
                    Ilce newIlce = new Ilce(Convert.ToInt16(ilceControl.id.Text), ilceControl.ad.Text, Convert.ToInt16(ilceControl.sehirId.Text));
                    BLL.Insert<Ilce>(newIlce);
                    dataGridView.DataSource = BLL.Get("Ilce");
                }
            } catch {
                MessageBox.Show("An error occured during insert");
            }
        }

        //Update
        private void button3_Click (object sender, EventArgs e)
        {
            try {
                if (cbSelect.SelectedIndex == 0) {
                    BLL bll = new BLL();
                    Sehir newSehir = new Sehir(Convert.ToInt16(sehirControl.id.Text), sehirControl.ad.Text);

                    bll.Update<Sehir>(newSehir);
                    dataGridView.DataSource = bll.Get("Sehir");
                } else if (cbSelect.SelectedItem.ToString() == "Kisi") {
                    BLL bll = new BLL();
                    Kisi newKisi = new Kisi(Convert.ToInt16(kisiControl.id.Text), kisiControl.ad.Text, kisiControl.soyad.Text, Convert.ToInt16(kisiControl.yas.Text), kisiControl.adres.Text, kisiControl.sehir.Text, kisiControl.ilce.Text);

                    bll.Update<Kisi>(newKisi);
                    dataGridView.DataSource = bll.Get("Kisi");
                } else if (cbSelect.SelectedItem.ToString() == "Ilce") {
                    BLL bll = new BLL();
                    Ilce newIlce = new Ilce(Convert.ToInt16(ilceControl.id.Text), ilceControl.ad.Text, Convert.ToInt16(ilceControl.sehirId.Text));
                    bll.Update<Ilce>(newIlce);
                    dataGridView.DataSource = bll.Get("Ilce");
                }
            } catch {
                MessageBox.Show("Error occured during update.");
            }
        }

        private void Form1_Load (object sender, EventArgs e)
        {
            this.Controls.Add(sehirControl);
            this.Controls.Add(kisiControl);
            this.Controls.Add(ilceControl);
            foreach (Control c in this.Controls)
                if (c is UserControl) c.Visible = false;
            cbSelect.SelectedIndex = 0;
        }

        //Combobox selection and filling the DataGridView
        private void comboBox1_SelectedValueChanged (object sender, EventArgs e)
        {
            foreach (Control u in this.Controls) {
                if (u.Name == cbSelect.SelectedItem + "Control") {
                    u.Visible = true;
                    u.Location = new Point(400, 0);
                    u.BringToFront();
                } else if (u is UserControl) u.Visible = false;
            }
            try {
                BLL bll = new BLL();
                dataGridView.DataSource = bll.Get(cbSelect.SelectedItem.ToString());
            } catch {
                MessageBox.Show("An error occured.");
            }
        }

        //To export Data to Excel file.
        private void button5_Click (object sender, EventArgs e)
        {
            updateFile();
        }

        //UPDATE the Excel file
        private void updateFile() {
            Excel.Application xlApplication = (Excel.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook xlWorkbook = (Excel.Workbook) xlApplication.ActiveWorkbook;
            Excel.Worksheet xlWorksheet = (Excel.Worksheet) xlWorkbook.ActiveSheet;
            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int i = 0; i < dataGridView.Rows.Count; i++) {
                for (int j = 0; j < dataGridView.Columns.Count; j++) {
                    xlWorksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value;
                }
            }
        }

        #region Deletion
        //DELETE by right click on row header. (not working properly)
        private void dataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e) {
            //if (e.Button == MouseButtons.Right) {
            //    dataGridView.MultiSelect = false;
            //    var hitTest = dataGridView.HitTest(e.X, e.Y);   //Location of mouse hit on datagridview
            //    dataGridView.ClearSelection();
            //    dataGridView.Rows[hitTest.RowIndex].Selected = true;    //Select the row which equals to hitTest's row index
            //    ContextMenu menu = new ContextMenu();
            //    MenuItem deleteMenuItem = new MenuItem("Delete");
            //    deleteMenuItem.Click += DeleteMenuItem_Click;   //Generate a Click event for deleteMenuItem MenuItem
            //    menu.MenuItems.Add(deleteMenuItem);
            //    menu.Show(dataGridView, new Point(e.X, e.Y));
            //}
        }

        //DELETE in Excel file
        private void DeleteMenuItem_Click(object sender, EventArgs e) {
            try {
                int deleteRow = dataGridView.SelectedRows[0].Index;
                dataGridView.Rows.RemoveAt(deleteRow);
                Excel.Application xlApplication = (Excel.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel.Workbook xlWorkbook = (Excel.Workbook) xlApplication.ActiveWorkbook;
                Excel.Worksheet xlWorksheet = (Excel.Worksheet) xlWorkbook.ActiveSheet;
                Excel.Range xlRange = xlWorksheet.UsedRange;
                ((Excel.Range) xlWorksheet.Rows[deleteRow + 2]).Delete(Type.Missing);
                updateFile();
            } catch {
                MessageBox.Show("Could not update the Excel file. Make sure the file is open.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //DELETE operation by right click on any point in grid.
        private void dataGridView_MouseClick(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right) {
                try {
                    dataGridView.MultiSelect = false;
                    var hitTest = dataGridView.HitTest(e.X, e.Y);   //Location of mouse hit on datagridview
                    dataGridView.CancelEdit();      //Cancel any editing cells
                    dataGridView.ClearSelection();  //Select no cell in grid
                    dataGridView.Rows[hitTest.RowIndex].Selected = true;    //Select the row which equals to hitTest's row index
                    ContextMenu menu = new ContextMenu();
                    MenuItem deleteMenuItem = new MenuItem("Delete");
                    deleteMenuItem.Click += DeleteMenuItem_Click;   //Generate a Click event for deleteMenuItem MenuItem
                    menu.MenuItems.Add(deleteMenuItem);
                    menu.Show(dataGridView, new Point(e.X, e.Y));
                } catch (ArgumentOutOfRangeException exception) {

                } catch (Exception exception) {
                    MessageBox.Show(exception.ToString());
                }

            }
        }
        #endregion

        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e) {
            dataGridView.BeginEdit(true);
        }
    }
}
