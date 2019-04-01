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
            crystalReportViewer1.RefreshReport();
        }
    }
}
