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
    public partial class LoginO : Form
    {
        public LoginO()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string constring = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Voluntaro.accdb";
            string cmdText = "select Count(*) from LoginOrg where Username=? and [Password]=?";
            using (OleDbConnection con = new OleDbConnection(constring))
            using (OleDbCommand cmd = new OleDbCommand(cmdText, con))
            {
                con.Open();
                cmd.Parameters.AddWithValue("@p1", tbusername.Text);
                cmd.Parameters.AddWithValue("@p2", tbpassword.Text);
                int result = (int)cmd.ExecuteScalar();
                if (result > 0)
                {

                    this.Hide();
                    Events ev = new Events();
                    ev.ShowDialog();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Check again your username/password!");
                }
                con.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Paginaprincipala p = new Paginaprincipala();
            p.ShowDialog();
            this.Close();
        }
    }
}
