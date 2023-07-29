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


namespace CreateSchedules
{
    /// <summary>
    /// Interaction logic for frmCreateSchedules.xaml
    /// </summary>
    public partial class frmCreateSchedules : Window
    {

        List<CheckBox> allCheckboxes = new List<CheckBox>();

        public frmCreateSchedules()
        {
            InitializeComponent();

            allCheckboxes.Add(chbIndex);
            allCheckboxes.Add(chbVeneer);
            allCheckboxes.Add(chbFloor);
            allCheckboxes.Add(chbFrame);
            allCheckboxes.Add(chbAttic);

            List<string> listElevations = new List<string> { "A", "B", "C", "D", "S", "T" };

            foreach (string elevation in listElevations)
            {
                cmbElevation.Items.Add(elevation);
            }

            cmbElevation.SelectedIndex = 0;

            List<string> listFloors = new List<string> { "1", "2", "3" };

            foreach (string floor in listFloors)
            {
                cmbFloors.Items.Add(floor);
            }

            cmbFloors.SelectedIndex = 0;
        }

        public string GetComboBoxElevationSelectedItem()
        {
            return cmbElevation.SelectedItem.ToString();
        }

        public int GetComboBoxFloorsSelectedItem()
        {
            return Int32.Parse(cmbFloors.SelectedItem.ToString());
        }

        public string GetGroup1()
        {
            if (rbBasement.IsChecked == true)
                return rbBasement.Content.ToString();
            else if (rbCrawlspace.IsChecked == true)
                return rbCrawlspace.Content.ToString();
            else
                return rbSlab.Content.ToString();
        }

        public bool GetCheckboxIndex()
        {
            if (chbIndex.IsChecked == true)
                return true;
            else
                return false;
        }

        public bool GetCheckboxVeneer()
        {
            if (chbVeneer.IsChecked == true)
                return true;
            else
                return false;
        }

        public bool GetCheckboxFloor()
        {
            if (chbFloor.IsChecked == true)
                return true;
            else
                return false;
        }

        public bool GetCheckboxFrame()
        {
            if (chbFrame.IsChecked == true)
                return true;
            else
                return false;
        }

        public bool GetCheckboxAttic()
        {
            if (chbAttic.IsChecked == true)
                return true;
            else
                return false;
        }

        public string GetGroup2()
        {
            if (rbSingle.IsChecked == true)
                return rbSingle.Content.ToString();
            else
                return rbMulti.Content.ToString();
        }

        private void btnAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (System.Windows.Controls.CheckBox cBox in allCheckboxes)
            {
                cBox.IsChecked = true;
            }
        }

        private void btnNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (System.Windows.Controls.CheckBox cBox in allCheckboxes)
            {
                cBox.IsChecked = false;
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}

