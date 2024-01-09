using System;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using MySql.Data.MySqlClient;

namespace Database2
{
    public partial class Form2 : Form
    {
        #region StartParametrs
        static string query;
        static string cmdString;
        static string pizzeriaSelected;
        static bool canContinue = true;
        string startDate;
        string endDate;
        string exceptedFormat;
        string itemType;
        string item;
        #endregion
        public Form2()
        {
            InitializeComponent();
        }
        private void mysqlConnection(string sqlCmd)
        {
            MySqlConnection cnn;
            string connectionPath = @"SERVER=server.com; PORT=3306; UID=Admin; PASSWORD=73iK91; DATABASE=pizzamondo;";
            cnn = new MySqlConnection(connectionPath);

            try
            {
                chart1.Titles.Clear();
                cnn.Open();
                DataTable dt = new DataTable();
                query = sqlCmd;
                MySqlDataAdapter MyDA = new MySqlDataAdapter(query, cnn);
                MyDA.Fill(dt);
                chart1.DataSource = dt;
                chart1.Series["Myynti"].XValueMember = "Tuote";
                chart1.Series["Myynti"].YValueMembers = "Määrä";
                chart1.Titles.Add("Tuotteiden myynti");
                chartType();
            }
            catch (Exception error)
            {
                MessageBox.Show(query);
                MessageBox.Show("Ongelma!!" + "\n" + error);
            }
            finally
            {
                cnn.Close();
            }
        }
        private string ItemType()
        {
            if (comboBox2.SelectedIndex > -1) 
            { 
                item = comboBox2.SelectedItem.ToString();
                if (comboBox2.SelectedIndex == 2) item = "lisatayte";
                canContinue = true; 
            }
            else
            {
                canContinue = false;
                MessageBox.Show("Valitse tuoteryhmä");
            }
            return item;
        }
        private void pizzeriaSelection()
        {
            itemType = ItemType();
            if (itemType == null) 
            { 
                canContinue = false;
                return; 
            }

            if (comboBox1.SelectedIndex == 2)
            {
                canContinue = true;
                pizzeriaSelected = "1 OR tilaukset.pizzeriaID = 2";
            }
            else if (comboBox1.SelectedIndex > -1)
            {
                canContinue = true;
                int selectedPizzeria = comboBox1.SelectedIndex + 1;
                pizzeriaSelected = selectedPizzeria.ToString();
            }
            else
            {
                canContinue = false;
                MessageBox.Show("Valitse Pizzeria");
            }
        }
        private void chartType()
        {
            switch (comboBox3.SelectedIndex)
            {
                case 0: chart1.Series["Myynti"].ChartType = SeriesChartType.Column; break;
                case 1: chart1.Series["Myynti"].ChartType = SeriesChartType.Bar; break;
                case 2: chart1.Series["Myynti"].ChartType = SeriesChartType.Pie; break;
                case 3: chart1.Series["Myynti"].ChartType = SeriesChartType.SplineArea; break;
                case 4: chart1.Series["Myynti"].ChartType = SeriesChartType.Radar; break;
                default: chart1.Series["Myynti"].ChartType = SeriesChartType.Column; break;
            }

        }

        #region TypesOfReport
        #region Day
        private void button2_Click(object sender, EventArgs e)
        {
            string date = monthCalendar1.SelectionStart.Date.ToString("yyyy-MM-dd");

            pizzeriaSelection(); if (canContinue == false) return;
            cmdString = ("SELECT " + itemType + ".nimi as `Tuote`, COUNT(*) AS `Määrä` FROM tilaukset JOIN " + itemType + " on tilaukset." + itemType + "ID = " + itemType + "." + itemType + "ID WHERE (tilaukset.pizzeriaID = " + pizzeriaSelected + ") AND tilauksen_pvm = " + "\"" + date + "\"" + " GROUP BY " + itemType + ".nimi;");
            mysqlConnection(cmdString);
        }
        #endregion

        #region Week
        private void button3_Click(object sender, EventArgs e)
        {
            pizzeriaSelection(); if (canContinue == false) return;
            cmdString = "SELECT " + itemType + ".nimi as `Tuote`, COUNT(*) AS `Määrä` FROM tilaukset JOIN " + itemType + " on tilaukset." + itemType + "ID = " + itemType + "." + itemType + "ID WHERE (tilaukset.pizzeriaID = " + pizzeriaSelected + ") AND tilauksen_pvm >= CURDATE() - INTERVAL WEEKDAY(CURDATE()) DAY - INTERVAL 1 WEEK AND tilauksen_pvm<CURDATE() -INTERVAL WEEKDAY(CURDATE()) DAY GROUP BY " + itemType + ".nimi ; ";
            mysqlConnection(cmdString);
        }
        #endregion

        #region Month
        private void button4_Click(object sender, EventArgs e)
        {
            pizzeriaSelection(); if (canContinue == false) return;
            cmdString = "SELECT " + itemType + ".nimi as `Tuote`, COUNT(*) AS `Määrä` FROM tilaukset JOIN " + itemType + " on tilaukset." + itemType + "ID = " + itemType + "." + itemType + "ID WHERE (tilaukset.pizzeriaID = " + pizzeriaSelected + ") AND tilauksen_pvm >= DATE_FORMAT(CURRENT_DATE - INTERVAL 1 MONTH, '%Y-%m-01') AND tilauksen_pvm < DATE_FORMAT(CURRENT_DATE, '%Y-%m-01') GROUP BY " + itemType + ".nimi";
            mysqlConnection(cmdString);
        }
        #endregion

        #region CustomDate
        private void button6_Click(object sender, EventArgs e)
        {
            startDate = textBox2.Text;
            endDate = textBox1.Text;
            exceptedFormat = "yyyy-MM-dd";
            if (DateTime.TryParseExact(startDate, exceptedFormat, null, System.Globalization.DateTimeStyles.None, out _) == false || DateTime.TryParseExact(endDate, exceptedFormat, null, System.Globalization.DateTimeStyles.None, out _) == false)
            {
                MessageBox.Show("Anna päivämäärä oikeassa muodossa (yyyy-mm-dd)");
                return;
            }
            pizzeriaSelection(); if (!canContinue) return;
            cmdString = "SELECT "+itemType+ ".nimi as `Tuote`, COUNT(*) AS `Määrä` FROM tilaukset JOIN " + itemType + " on tilaukset." + itemType + "ID = " + itemType + "." + itemType + "ID WHERE (tilaukset.pizzeriaID = " + pizzeriaSelected + ") AND tilauksen_pvm BETWEEN \"" + startDate + "\" AND \"" + endDate + "\" GROUP BY " + itemType + ".nimi;";
            mysqlConnection(cmdString);
        }
        
        #endregion
        #endregion
        private void button1_Click(object sender, EventArgs e) => this.Close();
    }
}
