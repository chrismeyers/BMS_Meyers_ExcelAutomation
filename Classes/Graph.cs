using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Graphs the means of all samples for each variable that is selected.
/// </summary>
/// <author>Chris Meyers</author>
namespace BMS_Meyers_ExcelAutomation.Classes {
    class Graph {
        String currentWs;
        String trimmedWsName;
        ArrayList variables;
        bool useSEM;
        Excel.Worksheet ws;
        Excel.Workbook activeWorkbook;

        /// <summary>
        /// The constructor for the class.
        /// </summary>
        /// <param name="ws">The "- FINAL" worksheet to be graphed.</param>
        /// <param name="variables">An ArrayList of all variables selected in the checkedListBox.</param>
        /// <param name="SEM">If true, graph with SEM error bars.</param>
        /// <author>Chris Meyers</author>
        public Graph(String ws, ArrayList variables, bool SEM) {
            currentWs = ws;
            this.variables = variables;
            useSEM = SEM;
        }

        /// <summary>
        /// Prepares a new "- PLOT" worksheet and and calls on helper methods
        /// to generate a graph.
        /// </summary>
        /// <author>Chris Meyers</author>
        public void plot() {
            activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            ws = Globals.ThisAddIn.Application.Sheets[currentWs];
            trimmedWsName = currentWs.Substring(0, currentWs.IndexOf(" "));

            Excel.Worksheet wsChart = (Excel.Worksheet)activeWorkbook.Worksheets.Add(); //Adds plot to a new ws
            try {
                //find any duplicate sheets and remove them 
                for (int i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++) {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Name.Contains(trimmedWsName + " - PLOT")) {
                        Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Delete();
                    }
                }

                //set the name of the new sheet
                wsChart.Name = trimmedWsName + " - PLOT";
            }
            catch {
                MessageBox.Show(wsChart.Name + " already exists please remove old before recalculating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsChart.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(0, 0, 500, 300);
            Excel.Chart chartPage = myChart.Chart;

            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

            // Set y-axis label
            Excel.Axis yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = trimmedWsName + " (% change from control)";
            yAxis.AxisTitle.Orientation = Excel.XlOrientation.xlUpward;

            // Store data from the "- FINAL" ws
            Dictionary<String, ArrayList> selectedVarData = parseData();

            // Create the plot on the "- PLOT" ws
            generatePlot(selectedVarData, myChart, chartPage);
        }

        /// <summary>
        /// Parses a "- FINAL" worsheet and stores the data to be used in graphing.
        /// </summary>
        /// <returns>A Dictionary containing Sample names, means and SEMs parsed from the selected 
        /// "- FINAL" worksheet</returns>
        /// <author>Chris Meyers</author>
        private Dictionary<String, ArrayList> parseData() {
            ArrayList sampleNames;
            ArrayList variableMeans;
            ArrayList variableSEMs;
            ArrayList allVariableData;
            Dictionary<String, ArrayList> allData = new Dictionary<String, ArrayList>();

            int numCols = ws.UsedRange.Columns.Count;
            int numRows = ws.UsedRange.Rows.Count;

            for(int vars = 2; vars < numRows; vars++){
                allVariableData = new ArrayList();
                String currentVar = ws.Cells[vars, 1].Value.ToString();
                if (variables.Contains(currentVar)) {
                    sampleNames = new ArrayList();
                    variableMeans = new ArrayList();
                    variableSEMs = new ArrayList();
                    for (int i = 2; i < numCols; i+=2) { 
                        sampleNames.Add(ws.Cells[1, i].Value.ToString());
                        variableMeans.Add(ws.Cells[vars, i].Value);
                        variableSEMs.Add(ws.Cells[vars, i + 1].Value);
                    }
                    allVariableData.Add(sampleNames);
                    allVariableData.Add(variableMeans);
                    allVariableData.Add(variableSEMs);
                }

                if (allVariableData.Count != 0) {
                    allData.Add(currentVar, allVariableData);
                }
            }

            return allData;
        }

        /// <summary>
        /// Uses the data from the "- FINAL" worksheet to produce a graph with the means of all samples for
        /// each variable that is selected.
        /// 
        /// If checkBox1 is checked, draw the standard error of the means.
        /// </summary>
        /// <param name="data">The data parsed from the "- FINAL" worksheet.</param>
        /// <param name="myChart">The current ChartObject.</param>
        /// <param name="chartPage">The current Chart.</param>
        /// <author>Chris Meyers</author>
        private void generatePlot(Dictionary<String, ArrayList> data, Excel.ChartObject myChart, Excel.Chart chartPage) {
            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)myChart.Chart.SeriesCollection();

            // Series labels print vertically (90 degrees).
            chartPage.Axes(Excel.XlAxisType.xlCategory).TickLabels.Orientation = 90;

            foreach (KeyValuePair<String, ArrayList> v in data) {
                Excel.Series currentGroup = seriesCollection.NewSeries();
                currentGroup.Name = trimmedWsName + " - " + v.Key;

                // Get sample names for current group.
                ArrayList sampleNames = new ArrayList();
                sampleNames = (ArrayList)data[v.Key][0];

                // Get mean values for current group.
                ArrayList means = new ArrayList();
                means = (ArrayList)data[v.Key][1];

                currentGroup.Values = means.ToArray(typeof(Double));
                currentGroup.XValues = sampleNames.ToArray(typeof(String));
            }

            // If checkBox1 is checked, draw error bars.
            if (useSEM) {
                int counter = 1;

                foreach (KeyValuePair<String, ArrayList> var in data) {
                    // Get SEMs for current group.
                    ArrayList SEMs = new ArrayList();
                    SEMs = (ArrayList)data[var.Key][2];

                    seriesCollection.Item(counter).ErrorBar(Excel.XlErrorBarDirection.xlX, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, 
                                                            Excel.XlErrorBarType.xlErrorBarTypeCustom, SEMs.ToArray(typeof(Double)), SEMs.ToArray(typeof(Double)));
                    counter++;
                }
            }
        }
    }
}
