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
        //Global variables
        private string _filePath;
        private int _totalMin;

        //NFL variables
        private int _qbMin;
        private int _rbMin;
        private int _wrMin;
        private int _teMin;
        private int _dstMin;
        
        //NBA variables
        private int _pgMin;
        private int _sgMin;
        private int _sfMin;
        private int _pfMin;
        private int _cMin;

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
            if ((bool) radioButton.IsChecked)
            {
                //Switch Case for NBA and NFL -- this one is NFL
                var matrixNFL = new NFLPlayerMatrix();
                SetNFLMinimums();
                var matrixReturn = matrixNFL.BuildPlayerList(_qbMin, _rbMin, _wrMin, _teMin, _dstMin, _filePath);
                var lineUp = new NFLLineUp();
                lineUp.BuildLineUp(matrixReturn, _totalMin, _filePath);
            } else
            {
                //--this one is for NBA
                var matrix = new NBAPlayerMatrix();
                SetNBAMinimums();
                var matrixReturnNBA = matrix.BuildPlayerList(_pgMin, _sgMin, _sfMin, _pfMin, _cMin, _filePath);
                var lineUpNBA = new NBALineUp();
                lineUpNBA.BuildLineUp(matrixReturnNBA, _totalMin, _filePath);
            }

        }

        private void radioButton_Checked(object sender, RoutedEventArgs e)
        {
            //this is my nFL radio button
        }

        private void radioButton1_Checked(object sender, RoutedEventArgs e)
        {
            //this is my NBA radio button
        }

        private void SetNFLMinimums()
        {
            _qbMin = Convert.ToInt32(textBox1.Text);
            _rbMin = Convert.ToInt32(textBox2.Text);
            _wrMin = Convert.ToInt32(textBox3.Text);
            _teMin = Convert.ToInt32(textBox4.Text);
            _dstMin = Convert.ToInt32(textBox5.Text);
            _totalMin = Convert.ToInt32(textBox6.Text);
        }

        private void SetNBAMinimums()
        {
            _pgMin = Convert.ToInt32(textBox1.Text);
            _sgMin = Convert.ToInt32(textBox2.Text);
            _sfMin = Convert.ToInt32(textBox3.Text);
            _pfMin = Convert.ToInt32(textBox4.Text);
            _cMin = Convert.ToInt32(textBox5.Text);
            _totalMin = Convert.ToInt32(textBox6.Text);
        }


    }
}
