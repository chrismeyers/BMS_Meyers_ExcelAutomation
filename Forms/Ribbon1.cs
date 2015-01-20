using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using BMS_Meyers_ExcelAutomation.Forms;

/// <summary>
/// Creates an add-in ribbon with options to format data and graph data.
/// </summary>
/// <author>Chris Meyers</author>
namespace BMS_Meyers_ExcelAutomation.Forms {
    public partial class Ribbon1 {
        /// <summary>
        /// The class constructor.
        /// </summary>
        /// <author>Chris Meyers</author>
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {
        }

        /// <summary>
        /// Instantiates and launches the format form (Form1).
        /// </summary>
        /// <author>Chris Meyers</author>
        private void format_Click(object sender, RibbonControlEventArgs e) {
            Form1 input = new Form1();
            input.Show();
        }

        /// <summary>
        /// Instantiates and launches the graph form (Form2).
        /// </summary>
        /// <author>Chris Meyers</author>
        private void graph_Click(object sender, RibbonControlEventArgs e) {
            Form2 input = new Form2();
            input.Show();
        }
    }
}
