using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// An input form used to select which variable(s) on which "- FINAL" worksheet to graph.
/// </summary>
/// <author>Chris Meyers</author>
namespace BMS_Meyers_ExcelAutomation.Forms {
    public partial class Form2 : Form {
        String type;

        /// <summary>
        /// The class constructor.
        /// </summary>
        /// <author>Chris Meyers</author>
        public Form2() {
            type = "graph";
            InitializeComponent();
            loadComboBoxes();
            loadCheckedListBox();
            checkedListBox1.CheckOnClick = true; // selects box on single-click
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
        /// Populates the checkedListBox with all the variable in the selected "- FINAL"
        /// worksheet.
        /// </summary>
        /// <author>Chris Meyers</author>
        public void loadCheckedListBox() {
            String selectedWorksheet = (String)comboBox1.SelectedItem;
            if (selectedWorksheet != null) { // If there is at least one "- FINAL" worksheet
                Excel.Worksheet ws = Globals.ThisAddIn.Application.Sheets[selectedWorksheet];

                checkedListBox1.Items.Clear();

                // Populate checked list box.
                int numVars = ws.UsedRange.Rows.Count;
                for (int i = 2; i <= numVars; i++) {
                    checkedListBox1.Items.Add(ws.Cells[i, 1].Value);
                }
            }
        }

        /// <summary>
        /// Stores the variables selectd from checkedListBox1 to an ArrayList.  Determines SEM functionality.
        /// Ensures at least one variable has be selected to be plotted.
        /// </summary>
        /// <author>Chris Meyers</author>
        private void button1_Click(object sender, EventArgs e) {
            bool SEM = false;
            String ws = (String)comboBox1.SelectedItem;
            ArrayList selectedVariables = new ArrayList();


            // Store selected variables
            for (int i = 0; i < checkedListBox1.Items.Count; i++) {
                if (checkedListBox1.GetItemChecked(i) == true) { 
                    selectedVariables.Add(checkedListBox1.Items[i].ToString());
                }
            }

            // Check to see if SEM is included
            if (checkBox1.Checked) {
                SEM = true;
            }

            if (selectedVariables.Count == 0) {
                MessageBox.Show("No variables selected. Please select at least 1 variable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else {
                Classes.Graph graph = new Classes.Graph(ws, selectedVariables, SEM);
                graph.plot();
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

        /// <summary>
        /// Populates the checkedListBox when the comboBox's value is changed.
        /// </summary>
        /// <author>Chris Meyers</author>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            loadCheckedListBox();
        }
    }
}
