using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormulariosAtos
{
    public partial class frm_CrystalReport : Form
    {
        public frm_CrystalReport()
        {
            InitializeComponent();           
        }

               public void GeraRelatorioBateria (int Valor)
        {
            
            if (Valor == 0)
            {
                CrystalReportViewer.ReportSource = LaudoBatText1;
                CrystalReportViewer.RefreshReport();
            }
            else
            {
                CrystalReportViewer.ReportSource = LaudoBatChart1;
                CrystalReportViewer.RefreshReport();
            }
        }
    }
}
