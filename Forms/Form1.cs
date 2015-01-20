using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMS_Meyers_ExcelAutomation.Classes;
using BMS_Meyers_ExcelAutomation.Forms;

/// <summary>
/// An input form used to select which form to format.
/// The number of experiments in this worksheet is also asked to aid in parsing.
/// </summary>
/// <author>Chris Meyers</author>
namespace BMS_Meyers_ExcelAutomation.Forms {
    public partial class Form1 : Form {
        String type;

        /// <summary>
        /// The class constructor.
        /// </summary>
        /// <author>Chris Meyers</author>
        public Form1() {
            type = "format";
            InitializeComponent();
            loadComboBoxes();
        }

        /// <summary>
        /// Populates the comboBox with applicable worksheets.
        /// </summary>
        /// <author>Chris Meyers</author>
        public void loadComboBoxes() {
            Forms.FormUtil util = new Forms.FormUtil(type);
            comboBox1.DataSource = util.loadComboBoxes();
        }

        /// <summary>
        /// Error checks the number of experiments input given by the user.  Instantiates formatData with the
        /// selected worksheet and the number of experiments.
        /// </summary>
        /// <author>Chris Meyers</author>
        private void button1_Click(object sender, EventArgs e) {
            String ws = (String)comboBox1.SelectedItem;
            int exper;

            bool isNumeric = int.TryParse(textBox1.Text, out exper);

            if (!isNumeric) {
                MessageBox.Show("Value for 'number of experiments' is invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Clear();
            }
            else {
                //MessageBox.Show(ws + "  " + exper);
                Classes.FormatData formatData = new Classes.FormatData(ws, exper);
                formatData.format();
                this.Close();
            }
        }

        /// <summary>
        /// Closes the form.
        /// </summary>
        /// <author>Chris Meyers</author>
        private void button2_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
