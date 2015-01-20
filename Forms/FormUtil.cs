using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Data;
using System.Drawing;

/// <summary>
/// Provides utilities commonly used in forms.
/// </summary>
/// <author>Chris Meyers</author>
namespace BMS_Meyers_ExcelAutomation.Forms {
    class FormUtil {
        String type;

        /// <summary>
        /// The class constructor.
        /// </summary>
        /// <param name="type">The type of form that is being used; either format or graph.</param>
        /// <author>Chris Meyers</author>
        public FormUtil(String type) {
            this.type = type;
        }

        /// <summary>
        /// Loads comboBoxes with worksheet names based on the type of form that was opened.
        /// --format forms: only return raw data sheets.
        /// --graph forms : only return "- FINAL" sheets.
        /// </summary>
        /// <returns>An ArrayList with all applicable worksheet names.</returns>
        /// <author>Chris Meyers</author>
        public ArrayList loadComboBoxes() {
            ArrayList allWorksheets = new ArrayList();

            if (type.Equals("format")) {
                for (int i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++) {
                    String currentSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Name;
                    // Only load raw ws's for format type
                    if (!currentSheet.Contains("FORMATTED") && !currentSheet.Contains("FINAL") && !currentSheet.Contains("PLOT")) {
                        allWorksheets.Add(currentSheet);
                    }
                }
            }
            else if(type.Equals("graph")){
                for (int i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++) {
                    String currentSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Name;
                    // Only load final ws's for graph type
                    if (currentSheet.Contains("FINAL")) {
                        allWorksheets.Add(currentSheet);
                    }
                }
            }

            return allWorksheets;
        }
    }
}
