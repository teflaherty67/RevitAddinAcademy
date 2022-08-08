using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAddinAcademy
{
    public partial class frmWallsFromLines : Form
    {
        public frmWallsFromLines(List<string> lineStyles, List<string> wallTypes)
        {
            InitializeComponent();

            foreach(string lineStyle in lineStyles)
            {
                this.cmbLineStyles.Items.Add(lineStyle);
            }

            foreach(string wallType in wallTypes)
            {
                this.cmbWallTypes.Items.Add(wallType);
            }

            this.cmbLineStyles.SelectedIndex = 0;
            this.cmbWallTypes.SelectedIndex = 0;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public string GetSelectedLineStyle()
        {
            return cmbLineStyles.SelectedItem.ToString();
        }

        public string GetSelectedWallType()
        {
            return cmbWallTypes.SelectedItem.ToString();
        }

        public double GetWallHeight()
        {
            double returnValue = 20;

            double.TryParse(tbxWallHeight.Text, out returnValue);

            return returnValue;
        }
    }
}
