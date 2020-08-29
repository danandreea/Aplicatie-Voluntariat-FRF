using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace frf
{
    public partial class Events : Form
    {
        public Events()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Paginaprincipala p = new Paginaprincipala();
            p.ShowDialog();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection conexiune = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Voluntaro.accdb");

            try
            {
                conexiune.Open();
                OleDbCommand comanda = new OleDbCommand();
                comanda.Connection = conexiune;
                comanda.CommandText = "SELECT MAX(ID) FROM Events";
                int cod = Convert.ToInt32(comanda.ExecuteScalar());
                comanda.CommandText = "INSERT INTO Events VALUES(?,?,?,?,?,?,?)";
                comanda.Parameters.Add("ID", OleDbType.Integer).Value = cod + 1;
                comanda.Parameters.Add("Nume", OleDbType.Char, 30).Value = tbNume.Text;
                comanda.Parameters.Add("Locatie", OleDbType.Char, 255).Value = tbLocatie.Text;
                comanda.Parameters.Add("DataI", OleDbType.DBDate).Value = dateTimePicker1.Value;
                comanda.Parameters.Add("DataF", OleDbType.DBDate).Value = dateTimePicker2.Value;
                comanda.Parameters.Add("Necesar", OleDbType.Integer).Value = tbNecesar.Text;
                comanda.Parameters.Add("Roluri", OleDbType.Char, 20).Value = tbRoluri.Text;
                comanda.ExecuteNonQuery();
                MessageBox.Show("Evenimentul a fost adaugat cu succes!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conexiune.Close();
                tbNume.Clear();
                tbLocatie.Clear();
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                tbNecesar.Clear();
                tbRoluri.Clear();
            }
        }
    }
}
