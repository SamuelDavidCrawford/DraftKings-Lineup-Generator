using System;
using System.Collections.Generic;
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
using System.IO;
using System.Windows.Forms;



namespace DraftKingsLineupGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _filePath;

        public MainWindow()
        {
            InitializeComponent();

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "CSV Files (.csv)|*.csv";
            openFileDialog1.FilterIndex = 1;

            openFileDialog1.Multiselect = false;
            
            DialogResult result = openFileDialog1.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                _filePath = openFileDialog1.FileName;
                try
                {
                    _filePath = openFileDialog1.FileName;
                    textBox.Text = _filePath;
                }
                catch (Exception)
                {
                    throw new ArgumentNullException("Cannot load CSV file.");
                }
            }

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
           this.Close();
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            var qbMin = Convert.ToInt32(qbMinSalary.Text);
            var rbMin = Convert.ToInt32(rbMinSalary.Text);
            var wrMin = Convert.ToInt32(wrMinSalary.Text);
            var teMin = Convert.ToInt32(teMinSalary.Text);
            var dstMin = Convert.ToInt32(dstMinSalary.Text);
            var totalMin = Convert.ToInt32(totalMinSalary.Text);


            //Builds player matrix (QB's, RB's, WR's, TE's, DST's, Flex's)
            var matrix = new PlayerMatrix();
            var matrixReturn = matrix.BuildPlayerList(qbMin, rbMin, wrMin, teMin, dstMin, _filePath);
            //Generates lineups
            var lineUp = new LineUp();
            lineUp.BuildLineUp(matrixReturn, totalMin, _filePath);
        }

    }
}
