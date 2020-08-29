using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft;
using Microsoft.Office.Interop.Excel;

namespace frf
{
    public partial class AdministratorPage : Form
    {
        string connString;

        public AdministratorPage()
        {
            InitializeComponent();
            connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Voluntaro.accdb";
            listavoluntari();
            listaevenimente();
        }

        private void listavoluntari()
        {
            OleDbConnection conexiune = new OleDbConnection(connString);
            try
            {
                listView1.Items.Clear();
                conexiune.Open();
                OleDbCommand comanda = new OleDbCommand();
                comanda.Connection = conexiune;
                comanda.CommandText = "SELECT * FROM Volunteers";
                OleDbDataReader reader = comanda.ExecuteReader();
                while (reader.Read())
                {
                    ListViewItem itm = new ListViewItem(reader["ID"].ToString());
                    itm.SubItems.Add(reader["Nume"].ToString());
                    itm.SubItems.Add(reader["Prenume"].ToString());
                    itm.SubItems.Add(reader["Data_Nasterii"].ToString());
                    itm.SubItems.Add(reader["E-mail"].ToString());
                    itm.SubItems.Add(reader["Parola"].ToString());
                    itm.SubItems.Add(reader["Telefon"].ToString());
                    itm.SubItems.Add(reader["Status"].ToString());
                    itm.SubItems.Add(reader["Score"].ToString());
                    listView1.Items.Add(itm);
                }

            
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conexiune.Close();
            }
        }
        private void listaevenimente()
        {
            OleDbConnection conexiune = new OleDbConnection(connString);
            try
            {
                listView2.Items.Clear();
                conexiune.Open();
                OleDbCommand comanda = new OleDbCommand();
                comanda.Connection = conexiune;
                comanda.CommandText = "SELECT * FROM Events";
                OleDbDataReader reader = comanda.ExecuteReader();
                while (reader.Read())
                {
                    ListViewItem itm = new ListViewItem(reader["ID"].ToString());
                    itm.SubItems.Add(reader["Nume"].ToString());
                    itm.SubItems.Add(reader["Locatie"].ToString());
                    itm.SubItems.Add(reader["DataI"].ToString());
                    itm.SubItems.Add(reader["DataF"].ToString());
                    itm.SubItems.Add(reader["Necesar"].ToString());
                    itm.SubItems.Add(reader["Roluri"].ToString());
                    listView2.Items.Add(itm);
                }

            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conexiune.Close();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Paginaprincipala p = new Paginaprincipala();
            p.ShowDialog();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog f = new OpenFileDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                listBox1.Items.Clear();

                List<string> lines = new List<string>();
                using (StreamReader r = new StreamReader(f.OpenFile()))
                {
                    string line;
                    while ((line = r.ReadLine()) != null)
                    {
                        listBox1.Items.Add(line);

                    }
                }
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            OleDbConnection conexiune = new OleDbConnection(connString);
            try
            {
                conexiune.Open();
                OleDbCommand comanda = new OleDbCommand();
                comanda.Connection = conexiune;
                foreach (ListViewItem itm in listView1.Items)
                {
                    if (itm.Checked==true)
                    {
                        int id = Convert.ToInt32(itm.SubItems[0].Text);
                        comanda.CommandText = "UPDATE Volunteers SET Status= -1 WHERE Id= "+id;
                        comanda.ExecuteNonQuery();
                        MessageBox.Show("Status-ul a fost modificat!");
                    }
                }

            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conexiune.Close();
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

            OleDbConnection conexiune = new OleDbConnection(connString);
            try
            {
                listView1.Items.Clear();
                conexiune.Open();
                OleDbCommand comanda = new OleDbCommand();
                comanda.Connection = conexiune;
                comanda.CommandText = "SELECT * FROM Volunteers";
                OleDbDataReader reader = comanda.ExecuteReader();
                while (reader.Read())
                {
                    ListViewItem itm = new ListViewItem(reader["ID"].ToString());
                    itm.SubItems.Add(reader["Nume"].ToString());
                    itm.SubItems.Add(reader["Prenume"].ToString());
                    itm.SubItems.Add(reader["Data_Nasterii"].ToString());
                    itm.SubItems.Add(reader["E-mail"].ToString());
                    itm.SubItems.Add(reader["Parola"].ToString());
                    itm.SubItems.Add(reader["Telefon"].ToString());
                    itm.SubItems.Add(reader["Status"].ToString());
                    itm.SubItems.Add(reader["Score"].ToString());
                    listView1.Items.Add(itm);
                }


            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conexiune.Close();
            }
        }

        private void topVolunteersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Top t = new Top();
            t.ShowDialog();
            this.Close();
        }

        private void backToMainPageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Workbook wb = xla.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)xla.ActiveSheet;
            int j = 1, i = 1;
            foreach (ListViewItem itm in listView1.Items)
            {
                ws.Cells[i, j] = itm.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in itm.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }
        }

        private void exportListaEvenimenteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Workbook wb = xla.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)xla.ActiveSheet;
            int j = 1, i = 1;
            foreach (ListViewItem itm in listView2.Items)
            {
                ws.Cells[i, j] = itm.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in itm.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }
        }
    }
    
}
