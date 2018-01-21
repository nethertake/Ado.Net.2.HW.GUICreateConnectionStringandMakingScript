using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace ADO.NET.HW2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string connectionStringType = "";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            if ((SelectingComboBox.SelectedIndex) == 1)
            {
                connectionStringType = "Provider=Microsoft.Jet.OLEDB.4.0;";
                Open.Visibility = Visibility.Visible;
                LabelNames.Content = "Data base path";
                Path.Text = "";
            }
            if ((SelectingComboBox.SelectedIndex) == 2)
            {
                connectionStringType = "providerName = \"System.Data.SqlClient\"; Initial Catalog = ;";
                LabelNames.Content = "Data base IP";
                Open.Visibility = Visibility.Hidden;
                Path.Text = "";
            }
        }


        private void ConnectButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {


                connectionStringType = connectionStringType + " User Id=" + TextBoxLogin.Text + "; Password=" + TextBoxPassword.Password + ";";
                ConnectionStringTextBlock.Text = "connectionString=" + connectionStringType;
                //string _connectionStringTest =
                //    "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=accessDb_00.mdb;User Id=admin;Password=;";
                string _connectionString = connectionStringType;
                OleDbConnection od = new OleDbConnection(_connectionString);
                od.Open();
                RealStatusBar.Content = od.State;
                MessageBox.Show("Connected");
                OleDbCommand cmd = new OleDbCommand("select * from newEquipment", od);
                cmd.ExecuteNonQuery();
                od.Close();
           


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Title = "Select a database";
            op.Filter = "Databases files|*.mdb;";
            if (op.ShowDialog() == true)
            {
                connectionStringType += " Data Source=" + op.FileName + ";";
                Path.Text = op.FileName;
                
            }




        }

        private void ExecuteButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string _connectionString = connectionStringType;
                OleDbConnection od = new OleDbConnection(_connectionString);
                od.Open();
               
                OleDbCommand cmd = new OleDbCommand(ScriptText.Text, od);
                cmd.ExecuteNonQuery();
                od.Close();
                RealStatusBar.Content = od.State;
                MessageStatusBar.Content = "Скрипт выполнен";
            }
            catch (Exception ex)
            {
                MessageStatusBar.Content = ex.Message;
            }


        }
    }
}
