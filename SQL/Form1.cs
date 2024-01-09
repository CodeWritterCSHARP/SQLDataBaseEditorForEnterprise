using System;
using System.Collections.Generic;
using System.Linq;
using MySql.Data.MySqlClient;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Database2
{
    public partial class Form1 : Form
    {
        #region StartValues
        //For Example
        private readonly string connectionPath = @"SERVER=server.com; PORT=3306; UID=Admin; PASSWORD=73iK91; DATABASE=pizzamondo;";
        private string type = null;
        private List<System.Windows.Forms.TextBox> textBoxes = new List<System.Windows.Forms.TextBox>();
        private List<System.Windows.Forms.Label> labels = new List<System.Windows.Forms.Label>();
        private int max = 0;
        #endregion

        public Form1()
        {
            InitializeComponent();
            InizializeLists();
        }
        private void button1_Click(object sender, EventArgs e) => Show();

        private void button2_Click(object sender, EventArgs e) => Add();

        private void button4_Click(object sender, EventArgs e) => Update();

        private void button5_Click(object sender, EventArgs e) => Delete();

        private void InizializeLists()
        {

            foreach (var item in flowLayoutPanel1.Controls.OfType<System.Windows.Forms.TextBox>()) textBoxes.Add(item);
            foreach (var item in flowLayoutPanel1.Controls.OfType<System.Windows.Forms.Label>()) labels.Add(item);
        }

        private new void Show()
        {
            if (String.IsNullOrEmpty(type)) return;
            MySqlConnection con = new MySqlConnection(connectionPath);
            try { con.Open(); } catch { MessageBox.Show("ei"); }
            string query = $"SELECT * FROM {type}";
            MySqlCommand cmd = new MySqlCommand(query, con);
            MySqlDataAdapter dataAdapter = new MySqlDataAdapter();
            dataAdapter.SelectCommand = cmd;
            System.Data.DataTable table = new System.Data.DataTable();
            dataAdapter.Fill(table);
            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = table;
            dataGridView1.DataSource = bindingSource;
            con.Close();
        }

        private void Add()
        {
            if (String.IsNullOrEmpty(type)) return;
            MySqlConnection con = new MySqlConnection(connectionPath);
            try { con.Open(); } catch { MessageBox.Show("ei"); }
            try
            {
                string query = null;
                List<string> values = new List<string>();
                for (int i = 0; i < textBoxes.Count; i++) values.Add(textBoxes[i].Text);

                switch (type)
                {
                    case "pizza":
                        query = "INSERT INTO pizza (pizzaID, nimi, täyte, hinta) VALUES('" + Convert.ToInt32(values[3]) + "', '" + values[0] + "', '" + values[1] + "', '" + Convert.ToInt32(values[2]) + "')";
                        break;

                    case "asiakas":
                        query = "INSERT INTO asiakas (asiakasID, Etunimi, Sukunimi, Puhelinnumero, Sahkoposti, Osoite) VALUES('" + Convert.ToInt32(values[0]) + "', '" + values[1] + "', '" + values[2] + "', '" + values[3] + "', '" + values[4] + "', '" + values[5] + "')";
                        break;

                    case "juoma":
                        query = "INSERT INTO juoma (juomaID, nimi, hinta) VALUES('" + Convert.ToInt32(values[0]) + "', '" + values[1] + "', '" + Convert.ToInt32(values[2]) + "')";
                        break;

                    case "juomarivi":
                        query = "INSERT INTO juomarivi (juomariviID, juomaID) VALUES('" + Convert.ToInt32(values[0]) + "', '" + Convert.ToInt32(values[1]) + "')";
                        break;

                    case "kayttajaryhma":
                        query = "INSERT INTO kayttajaryhma (kayttajaryhmaID, KayttajaryhmaNimi) VALUES('" + Convert.ToInt32(values[0]) + "', '" + values[1] + "')";
                        break;

                    case "kayttajat":
                        query = "INSERT INTO kayttajat (kayttajatID, asiakasID, KayttajaryhmaID, kayttajanimi, salasana) VALUES('" + Convert.ToInt32(values[0]) + "', '" + Convert.ToInt32(values[1]) + "', '" + Convert.ToInt32(values[2]) + "', '" + values[3] + "', '" + values[4] + "')";
                        break;

                    case "lisatayte":
                        query = "INSERT INTO lisatayte (lisatayteID, nimi, hinta) VALUES('" + Convert.ToInt32(values[0]) + "', '" + values[1] + "', '" + Convert.ToInt32(values[2]) + "')";
                        break;

                    case "lisatayterivi":
                        query = "INSERT INTO lisatayterivi (lisatayteriviID, lisatayteID) VALUES('" + Convert.ToInt32(values[0]) + "', '" + Convert.ToInt32(values[1]) + "')";
                        break;

                    case "Pizzeria":
                        query = "INSERT INTO Pizzeria (PizzeriaID, Osoite, Omistaja) VALUES('" + Convert.ToInt32(values[0]) + "', '" + values[1] + "', '" + values[2] + "')";
                        break;

                    case "tilaukset":
                        query = "INSERT INTO tilaukset (tilauksetID, asiakasID, pizzeriaID, toimitusosoite, tilauksen_pvm) VALUES('" + Convert.ToInt32(values[0]) + "', '" + Convert.ToInt32(values[1]) + "', '" + Convert.ToInt32(values[2]) + "', '" + values[3] + "', NOW())";
                        break;

                    case "tilausrivi":
                        query = "INSERT INTO tilausrivi (tilausriviID, tilausID, pizzaID, lisatayteriviID, juomariviID) VALUES('" + Convert.ToInt32(values[0]) + "', '" + Convert.ToInt32(values[1]) + "', '" + Convert.ToInt32(values[2]) + "', '" + Convert.ToInt32(values[3]) + "', '" + Convert.ToInt32(values[4]) + "')";
                        break;

                    default: break;
                }
                MySqlCommand cmd = new MySqlCommand(query, con);
                if (cmd.ExecuteNonQuery() == 1) MessageBox.Show("was added"); else MessageBox.Show("wasnt added");
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            con.Close();
            Show();
        }

        private new void Update()
        {
            try
            {
                SearchMax();

                List<string> cells = new List<string>();
                for (int j = 0; j < dataGridView1.ColumnCount; j++) cells.Add(dataGridView1.Columns[j].Name.ToString());

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    MySqlConnection con = new MySqlConnection(connectionPath);
                    try { con.Open(); } catch { MessageBox.Show("ei"); }
                    string query = "";

                    if (i <= max - 1)
                    {
                        switch (type)
                        {
                            case "pizza":
                                query = "UPDATE pizza SET nimi= '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() + "', hinta= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["hinta"].Value.ToString()) + 
                                    "', täyte= '" + dataGridView1.Rows[i].Cells["täyte"].Value.ToString() + "' WHERE pizzaID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["pizzaID"].Value.ToString()) + "'";
                                break;

                            case "asiakas":
                                query = "UPDATE asiakas SET Etunimi= '" + dataGridView1.Rows[i].Cells["Etunimi"].Value.ToString() + "', Sukunimi= '" + dataGridView1.Rows[i].Cells["Sukunimi"].Value.ToString() +
                                    "', Puhelinnumero= '" + dataGridView1.Rows[i].Cells["täyte"].Value.ToString() + "', Sahkoposti= '" + dataGridView1.Rows[i].Cells["Sahkoposti"].Value.ToString() +
                                    "', Osoite= '" + dataGridView1.Rows[i].Cells["Osoite"].Value.ToString() + "' WHERE asiakasID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["asiakasID"].Value.ToString()) + "'";
                                break;

                            case "juoma":
                                query = "UPDATE juoma SET nimi= '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() + "', hinta= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["hinta"].Value.ToString()) + "' WHERE juomaID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomaID"].Value.ToString()) + "'";
                                break;

                            case "juomarivi":
                                query = "UPDATE juomarivi SET juomaID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomaID"].Value.ToString()) + "' WHERE juomariviID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomariviID"].Value.ToString()) + "'";
                                break;

                            case "kayttajaryhma":
                                query = "UPDATE kayttajaryhma SET KayttajaryhmaNimi= '" + dataGridView1.Rows[i].Cells["KayttajaryhmaNimi"].Value.ToString() + "' WHERE kayttajaryhmaID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["kayttajaryhmaID"].Value.ToString()) + "'";
                                break;

                            case "kayttajat":
                                query = "UPDATE kayttajat SET asiakasID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["asiakasID"].Value.ToString()) + "', KayttajaryhmaID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["KayttajaryhmaID"].Value.ToString()) +
                                    "', kayttajanimi= '" + dataGridView1.Rows[i].Cells["kayttajanimi"].Value.ToString() + "', salasana= '" + dataGridView1.Rows[i].Cells["salasana"].Value.ToString() + "' WHERE kayttajatID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["kayttajatID"].Value.ToString()) + "'";
                                break;

                            case "lisatayte":
                                query = "UPDATE lisatayte SET nimi= '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() + "', hinta= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["hinta"].Value.ToString()) + "' WHERE lisatayteID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["lisatayteID"].Value.ToString()) + "'";       
                                break;

                            case "lisatayterivi":
                                query = "UPDATE lisatayterivi SET lisatayteID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomaID"].Value.ToString()) + "' WHERE lisatayteriviID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["lisatayteriviID"].Value.ToString()) + "'";
                                break;

                            case "Pizzeria":
                                query = "UPDATE Pizzeria SET Osoite= '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() + "', Omistaja= '" + dataGridView1.Rows[i].Cells["Omistaja"].Value.ToString() + "' WHERE PizzeriaID= '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["PizzeriaID"].Value.ToString()) + "'";
                                break;

                            default: break;
                        }
                        MySqlCommand cmd = new MySqlCommand(query, con);
                        if (cmd.ExecuteNonQuery() == 1) { } else { }
                    }
                    else
                    {
                        switch (type)
                        {
                            case "pizza":
                                query = "INSERT INTO pizza (pizzaID, nimi, täyte, hinta) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["pizzaID"].Value.ToString()) + "', '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() +
                                    "', '" + dataGridView1.Rows[i].Cells["täyte"].Value.ToString() + "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["hinta"].Value.ToString()) + "')";
                                break;

                            case "juoma":
                                query = "INSERT INTO juoma (juomaID, nimi, hinta) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomaID"].Value.ToString()) + "', '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() + "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["hinta"].Value.ToString()) + "')";
                                break;

                            case "juomarivi":
                                query = "INSERT INTO juomarivi (juomariviID, juomaID) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomariviID"].Value.ToString()) + "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["juomaID"].Value.ToString()) + "')";
                                break;

                            case "kayttajaryhma":
                                query = "INSERT INTO kayttajaryhma (kayttajaryhmaID, KayttajaryhmaNimi) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["kayttajaryhmaID"].Value.ToString()) + "', '" + dataGridView1.Rows[i].Cells["KayttajaryhmaNimi"].Value.ToString() + "')";
                                break;

                            case "kayttajat":
                                query = "INSERT INTO kayttajat (kayttajatID, asiakasID, KayttajaryhmaID, kayttajanimi, salasana) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["kayttajatID"].Value.ToString()) + "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["asiakasID"].Value.ToString()) +
                                    "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["KayttajaryhmaID"].Value.ToString()) + "', '" + dataGridView1.Rows[i].Cells["kayttajanimi"].Value.ToString() + "', '" + dataGridView1.Rows[i].Cells["salasana"].Value.ToString() + "')";
                                break;

                            case "lisatayte":
                                query = "INSERT INTO lisatayte (lisatayteID, nimi, hinta) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["lisatayteID"].Value.ToString()) + "', '" + dataGridView1.Rows[i].Cells["nimi"].Value.ToString() + "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["hinta"].Value.ToString()) + "')";
                                break;

                            case "lisatayterivi":
                                query = "INSERT INTO lisatayterivi (lisatayteriviID, lisatayteID) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["lisatayteriviID"].Value.ToString()) + "', '" + Convert.ToInt32(dataGridView1.Rows[i].Cells["lisatayteID"].Value.ToString()) + "')";
                                break;

                            case "Pizzeria":
                                query = "INSERT INTO Pizzeria (PizzeriaID, Osoite, Omistaja) VALUES('" + Convert.ToInt32(dataGridView1.Rows[i].Cells["PizzeriaID"].Value.ToString()) + "', '" + dataGridView1.Rows[i].Cells["Osoite"].Value.ToString() + "', '" + dataGridView1.Rows[i].Cells["Omistaja"].Value.ToString() + "')";
                                break;

                            default: break;
                        }
                        MySqlCommand cmd = new MySqlCommand(query, con);
                        if (cmd.ExecuteNonQuery() == 1) { } else { }
                    }
                    con.Close();
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            Show();
        }

        private void SearchMax()
        {
            MySqlConnection con = new MySqlConnection(connectionPath);
            con.Open();
            string query = $"SELECT COUNT(*) FROM {type}";
            MySqlCommand cmd = new MySqlCommand(query, con);
            max = Convert.ToInt32(cmd.ExecuteScalar());
            con.Close();
        }

        private void Delete()
        {
            MySqlConnection con = new MySqlConnection(connectionPath);
            try { con.Open(); } catch { MessageBox.Show("ei"); }
            try
            {
                int id = Convert.ToInt32(textBox5.Text);
                string query = $"DELETE FROM {type} WHERE {type}ID= '" + id + "'";
                MySqlCommand cmd = new MySqlCommand(query, con);
                if (cmd.ExecuteNonQuery() == 1) MessageBox.Show("was deleted"); else MessageBox.Show("wasnt deleted");
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            con.Close();
            Show();
        }

        #region SwitchTableType
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) 
        { 
            type = comboBox1.Text;
            switch (type)
            {
                case "pizza":
                    labels[0].Text = "nimi"; labels[1].Text = "täyte"; 
                    labels[2].Text = "hinta"; labels[3].Text = "pizzaID";
                    SwitchType(4);
                break;

                case "Pizzeria":
                    labels[0].Text = "PizzeriaID"; labels[1].Text = "Osoite"; labels[2].Text = "Omistaja";
                    SwitchType(3);
                    break;

                case "asiakas":
                    labels[0].Text = "asiakasID"; labels[1].Text = "Etunimi"; labels[2].Text = "Sukunimi"; 
                    labels[3].Text = "Puh nmro"; labels[4].Text = "e-posti"; labels[5].Text = "Osoite";
                    SwitchType(6);
                    break;

                case "kayttajaryhma":
                    labels[0].Text = "RyhmäID"; labels[1].Text = "RyhmäNimi";
                    SwitchType(2);
                    break;

                case "kayttajat":
                    labels[0].Text = "KäyttäjäID"; labels[1].Text = "asiakasID"; labels[2].Text = "ryhmäID"; 
                    labels[3].Text = "Nimi"; labels[4].Text = "salasana";
                    SwitchType(5);
                    break;

                case "lisatayte":
                    labels[0].Text = "täyteID"; labels[1].Text = "nimi"; labels[2].Text = "hinta";
                    SwitchType(3);
                    break;

                case "juoma":
                    labels[0].Text = "juomaID"; labels[1].Text = "nimi"; labels[2].Text = "hinta";
                    SwitchType(3);
                    break;

                case "juomarivi":
                    labels[0].Text = "juomariviID"; labels[1].Text = "juomaID";
                    SwitchType(2);
                    break;

                case "lisatayterivi":
                    labels[0].Text = "riviID"; labels[1].Text = "täyteID";
                    SwitchType(2);
                    break;

                case "tilaukset":
                    labels[0].Text = "tilausID"; labels[1].Text = "asiakasID"; 
                    labels[2].Text = "pizzeriID"; labels[3].Text = "toimitusOs";
                    SwitchType(4);
                    break;

                case "tilausrivi":
                    labels[0].Text = "tilausriviID"; labels[1].Text = "tilausID"; labels[2].Text = "pizzaID"; 
                    labels[3].Text = "lisatriiviID"; labels[4].Text = "juomariviID";
                    SwitchType(5);
                    break;

                default: break;
            }
            Show();
        }

        private void SwitchType(int border)
        {
            for (int i = 0; i < textBoxes.Count; i++)
            {
                if (i >= border)
                {
                    textBoxes[i].Visible = false; textBoxes[i].Enabled = false;
                    labels[i].Visible = false; labels[i].Enabled = false;
                }
                else
                {
                    textBoxes[i].Visible = true; textBoxes[i].Enabled = true;
                    labels[i].Visible = true; labels[i].Enabled = true;
                }
            }
        }
        #endregion

        #region ExportToExcel
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 1)
            {
                try
                {
                    _Application app = new Microsoft.Office.Interop.Excel.Application();
                    _Workbook workbook = app.Workbooks.Add(Type.Missing);
                    _Worksheet worksheet = null;
                    app.Visible = false;
                    worksheet = workbook.Sheets["Taul1"];
                    for (int i = 1; i < dataGridView1.ColumnCount + 1; i++) worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            if (dataGridView1.Rows[i].Cells[j].Value != null) worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            else worksheet.Cells[i + 2, j + 1] = "";
                        }
                    }
                    try { workbook.SaveAs($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/{textBox8.Text}.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing); MessageBox.Show("Saved"); }
                    catch (Exception ex1) { MessageBox.Show(ex1.Message); }
                    finally { app.Quit(); }
                }
                catch (Exception ex1) { MessageBox.Show(ex1.Message); }
            }
            else { MessageBox.Show("Taulukossa ei oo mitään"); }
        }
        #endregion

        #region SecondFormActivation
        private void button3_Click(object sender, EventArgs e)
        {
            if (System.Windows.Forms.Application.OpenForms.Count < 2) 
            {
                Form2 form2 = new Form2(); 
                form2.Show();
            }
        }
        #endregion
    }
}