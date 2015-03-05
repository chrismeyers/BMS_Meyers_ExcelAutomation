using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Parses an initial worksheet and creates two new worksheets based on the intial worksheet's
/// data.  
/// 
/// The first new worksheet entitled "- FORMATTED" displays each variable for each sample for
/// each experiment and it's associated value.  Statistics for each of these samples are also 
/// calculated and printed under the sample.
/// 
/// The second new worksheet entitled "- FINAL" condenses the calculated stats to the mean and
/// SEM for each variable and each sample (excluding the controls).  This worksheet is used in
/// graphing.
/// </summary>
/// <author>Chris Meyers</author>
namespace BMS_Meyers_ExcelAutomation.Classes {
    public class FormatData {
        Excel.Worksheet ws;
        Excel.Worksheet formattedWs;
        Excel.Worksheet finalWs;
        Excel.Workbook activeWorkbook;
        object misValue = System.Reflection.Missing.Value;
        String currentWs;
        int numExperiments;
        int numVariables;
        int numSamples;
        ArrayList variableNames;
        ArrayList sampleNames;
        ArrayList means;
        ArrayList SEMs;
        List<List<Dictionary<String, Dictionary<String, Double>>>> experimentsData;

        /// <summary>
        /// The constructor for the class.
        /// </summary>
        /// <param name="selectedWs">The worksheet that was selcted in the format form.</param>
        /// <param name="experiments">The number of experiments recorded in selectedWS.</param>
        /// <author>Chris Meyers</author>
        public FormatData(String selectedWs, int experiments) {
            currentWs = selectedWs;
            numExperiments = experiments;
            numVariables = -1;
            numSamples = -1;
            means = new ArrayList();
            SEMs = new ArrayList();
            sampleNames = new ArrayList();
            variableNames = new ArrayList();
            experimentsData = new List<List<Dictionary<String, Dictionary<String, Double>>>>();
        }

        /// <summary>
        /// The primary method of the class.  Creates the two new worksheets and calls on helper methods
        /// to format these worksheets.
        /// </summary>
        /// <author>Chris Meyers</author>
        public void format() {
            // Get current spreadsheet
            ws = Globals.ThisAddIn.Application.Sheets[currentWs]; // Sets the current ws to the ws selected in Form1
            activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            // Parse all the data and store in List experimentsData.
            experimentsData = parseData();

            // Add new formatted worksheet
            formattedWs = activeWorkbook.Worksheets.Add();
            try {
                //find any duplicate sheets and remove them 
                for (int i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++) {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Name.Contains(ws.Name + " - FORMATTED")) {
                        Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Delete();
                    }
                }

                //set the name of the new sheet
                formattedWs.Name = ws.Name + " - FORMATTED";
            }
            catch {
                MessageBox.Show(formattedWs.Name + " already exists please remove old before recalculating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Add new final worksheet
            finalWs = activeWorkbook.Worksheets.Add();
            try {
                //find any duplicate sheets and remove them 
                for (int i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++) {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Name.Contains(ws.Name + " - FINAL")) {
                        Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[i].Delete();
                    }
                }

                //set the name of the new sheet
                finalWs.Name = ws.Name + " - FINAL";
            }
            catch {
                MessageBox.Show(finalWs.Name + " already exists please remove old before recalculating", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            // Print the data to a formatted ws
            makeFormattedWorksheet(experimentsData);

            // Expand columns to fit data
            formattedWs.Columns.AutoFit();

            // Perform statistics on the formatted ws
            calculateStats();

            // Print final stats needed to graph on the final ws
            makeFinalWorksheet(experimentsData);

            // Expand columns to fit data
            finalWs.Columns.AutoFit();
        }

        /// <summary>
        /// Parses a raw worksheet and stores all the data in a single list.
        /// </summary>
        /// <returns>A list containing the data for all experiments.</returns>
        /// <author>Chris Meyers</author>
        private List<List<Dictionary<String, Dictionary<String, Double>>>> parseData() {
            List<List<Dictionary<String, Dictionary<String, Double>>>> experiments = new List<List<Dictionary<String, Dictionary<String, Double>>>>();
            List<Dictionary<String, Dictionary<String, Double>>> experiment;
            Dictionary<String, Dictionary<String, Double>> sample;
            Dictionary<String, Double> sampleDictionary;

            int i = 1;
            int j = 1;
            int counter = 0;

            for (int exp = 1; exp <= numExperiments; exp++) { // experiments loop
                j = getTopRowForExperiment(exp);
                experiment = new List<Dictionary<String, Dictionary<String, Double>>>();

                while (ws.Cells[j, i].Value != null) { // experiment loop
                    String sampleName = "";
                    sampleDictionary = new Dictionary<String, Double>();
                    while (ws.Cells[j, i + 1].Value != null) {  // sample loop
                        // Convert to double. Disregard fist line (Transient #1).
                        // Set VALUE to 0.0 if equal to "Error"
                        if (counter != 0) {
                            String currentVar = ws.Cells[j, i + 1].Value.ToString();
                            Double currentVarConverted;

                            if (currentVar.Equals("Error")) {
                                currentVarConverted = 0.0;
                            }
                            else {
                                currentVarConverted = Convert.ToDouble(currentVar);
                            }
                            //MessageBox.Show(ws.Cells[j, i].Value.ToString() + " " + ws.Cells[j, i + 1].Value.ToString());

                            // Add entry to current sampleDictionary.
                            sampleDictionary.Add(ws.Cells[j, i].Value.ToString(), currentVarConverted);
                            counter++;
                        }
                        else {
                            //MessageBox.Show("SAMPLE NAME: " +ws.Cells[j, i].Value.ToString());
                            // Store the current sample name.
                            sampleName = ws.Cells[j, i].Value.ToString();
                            counter++;
                        }

                        j++;
                    }

                    // Add a new sample.
                    sample = new Dictionary<String, Dictionary<String, Double>>();
                    sample.Add(sampleName, sampleDictionary);

                    // Add sample to experiment.
                    experiment.Add(sample);

                    i = i + 3; //Move over to next sample in experiment
                    j = getTopRowForExperiment(exp);
                    counter = 0;
                    // Sets number of samples per experiment to max of all experiments
                    if (experiment.Count > numSamples) {
                        numSamples = experiment.Count;
                    }
                    
                }
                // Reset position for next experiment
                i = 1;

                // Add experiment to experiments
                experiments.Add(experiment);
            }

            return experiments;
        }

        /// <summary>
        /// Calculates and returns the starting row position for an experiment.
        /// Based on the number of variables in a sample and the current experiment.
        /// 
        /// !!!May need to be re-worked if the number of variables change!!!
        /// 
        /// </summary>
        /// <param name="exp">The current experiment parseData is processing.</param>
        /// <returns>The top row for the current experiment.</returns>
        /// <author>Chris Meyers</author>
        private int getTopRowForExperiment(int exp) {
            if (exp == 1) {
                return 1;
            }
            else {
                // Reset j to first row of experiment.
                // MAY NEED TO BE CHANGED IF THE # OF VARIABLES CHANGE.
                int header = 1;
                int numberOfVariables = 19;
                int spaceBtwExperiments = 2;
                int widthOfAnExperiment = header + numberOfVariables + spaceBtwExperiments;
                return ((exp - 1) * widthOfAnExperiment) + 1;
            }
        }

        /// <summary>
        /// Creates the "- FORMATTED" worksheet with each variable for each sample for each experiment.
        /// </summary>
        /// <param name="expData">The list created by parseData that contains all the data from
        /// the raw worksheet.</param>
        /// <author>Chris Meyers</author>
        private void makeFormattedWorksheet(List<List<Dictionary<String, Dictionary<String, Double>>>> expData) {
            int expNum = 0;
            int i = 1;
            int j = 1;
            ArrayList variablesInExperiments = new ArrayList();
            ArrayList samplesInExperiment = new ArrayList();

            for (int expIndex = 0; expIndex < numSamples; expIndex++) {
                foreach (List<Dictionary<String, Dictionary<String, Double>>> e in expData) { // experiment loop
                    Dictionary<String, Dictionary<String, Double>> sample = new Dictionary<String, Dictionary<String, Double>>();
                    sample = e[expIndex];
                    // Save the number of samples in each experiment
                    samplesInExperiment.Add(e.Count);
                    foreach (KeyValuePair<String, Dictionary<String, Double>> s in sample) { // sample loop
                        String sampleName = s.Key;
                        Dictionary<String, Double> data = new Dictionary<String, Double>();
                        data = s.Value;

                        // Adds the number of variables in each sample to an ArrayList
                        variablesInExperiments.Add(data.Count);

                        // Add unique sample name to global sample name ArrayList
                        if (!sampleNames.Contains(sampleName)) {
                            sampleNames.Add(sampleName);
                        }

                        // Print out data for all variables
                        int counter = 0;
                        int row = 0;
                        int offset = 0;
                        foreach (KeyValuePair<String, Double> d in data) {
                            // Add unique variable name to global variable name ArrayList
                            if (!variableNames.Contains(d.Key)) {
                                variableNames.Add(d.Key);
                            }

                            if (counter == 0) {
                                row = i + (numExperiments * counter + 1) - 1;
                            }
                            else {
                                // Change (numExperiments + N) to change space between variables.
                                row = i + offset + (counter * (numExperiments + 5));
                            }
                            print(data, row, j, expNum, sampleName, d.Key);
                            counter++;
                            offset++;
                        }

                        // Print out data for Specific variables
                        //print(data, i, j, counter, expNum, sampleName, "bl");
                        //print(data, i + numExperiments + 1, j, counter, expNum, sampleName, "bl%peak h");

                        i++;
                    }
                    expNum++;
                }

                i = 1; // Starts next experiment at top

                if (expIndex == 0) { // Adds space between experiments.
                    j += 3;
                }
                else {
                    j += 4;
                }
            }

            // Determine the max number of variables, set equal to numVariables global for stat calculation purposes.
            variablesInExperiments.Sort();
            variablesInExperiments.Reverse();
            numVariables = (int)variablesInExperiments[0];
        }

        /// <summary>
        /// Prints data to the "- FORMATTED" worksheet.
        /// </summary>
        /// <param name="data">The Dictionary that contains the current sample Variable-Value pairs.</param>
        /// <param name="i">The current row.</param>
        /// <param name="j">The current column.</param>
        /// <param name="expNum">The current experiment number.</param>
        /// <param name="sampleName">The name of the current sample to be printed.</param>
        /// <param name="variable">The variable to be printed.</param>
        /// <author>Chris Meyers</author>
        private void print(Dictionary<String, Double> data, int i, int j, int expNum, String sampleName, String variable) {
            if (data.Count == 0) {
                formattedWs.Cells[i, j] = "";
            }
            else {
                foreach (String key in data.Keys) { // data loop
                    Double value = data[key];
                    if (key.Equals(variable)) { // variable loop
                        // Print sample : variable
                        formattedWs.Cells[i, j] = sampleName + " : " + key;
                        // Print value next to 'sample : variable'
                        formattedWs.Cells[i, j + 1] = value;
                    }
                }
            }
        }

        /// <summary>
        /// Calculates and prints statistics for each variable for each sample for each experiment on the "- FORMATTED"
        /// worksheet.
        /// </summary>
        /// <author>Chris Meyers</author>
        private void calculateStats() {
            int startPoint = 2;
            int locationOfAOffset = 4;
            int locationOfBStatic = 2;
            int locationOfCOffset = 4;
            int iOffset = numExperiments + 6;

            for (int i = 0; i < numVariables; i++) {
                for (int j = 1; j <= numSamples - 1; j++) {
                    formattedWs.Cells[(i * iOffset) + numExperiments + 1, (startPoint - 1) + (locationOfAOffset * j)] = "mean=";
                    formattedWs.Cells[(i * iOffset) + numExperiments + 2, (startPoint - 1) + (locationOfAOffset * j)] = "n=";
                    formattedWs.Cells[(i * iOffset) + numExperiments + 3, (startPoint - 1) + (locationOfAOffset * j)] = "StDev=";
                    formattedWs.Cells[(i * iOffset) + numExperiments + 4, (startPoint - 1) + (locationOfAOffset * j)] = "SEM=";

                    ArrayList cValues = new ArrayList();
                    for (int e = 1; e <= numExperiments; e++) {
                        if (formattedWs.Cells[e + (i * iOffset), (startPoint - 1) + (locationOfAOffset * j)].Value != null) {
                            Double a = formattedWs.Cells[e + (i * iOffset), (startPoint - 1) + (locationOfAOffset * j)].Value;
                            Double b = formattedWs.Cells[e + (i * iOffset), locationOfBStatic].Value;
                            // Print out col "c"
                            Double cVal = ((a - b) / b) * 100;
                            cValues.Add(cVal);
                            formattedWs.Cells[e + (i * iOffset), startPoint + (locationOfCOffset * j)] = cVal;
                        }
                    }

                    int cValN = cValues.Count;
                    Double cValSqrtN = Math.Sqrt(cValN);
                    Double cValSum = 0.0;
                    Double cValMean = 0.0;
                    Double cValStDev = 0.0;

                    // Calculate sum of c values
                    for (int val = 0; val < cValN; val++) {
                        cValSum += (Double)cValues[val];
                    }

                    // Calculate mean of c values
                    cValMean = cValSum / cValN;

                    // Calculate StDev of c values
                    Double StDevNumeratorSum = 0.0;
                    for (int a = 0; a < cValN; a++) {
                        Double m = Math.Pow((Double)cValues[a] - cValMean, 2.0);
                        StDevNumeratorSum += m;
                    }
                    cValStDev = Math.Sqrt(StDevNumeratorSum / (cValN - 1));

                    // Calculate SEM of c values
                    Double cSEM = cValStDev / cValSqrtN;

                    // Print calculated Mean
                    formattedWs.Cells[(i * iOffset) + numExperiments + 1, startPoint + (locationOfAOffset * j)] = cValMean;
                    means.Add(cValMean);
                    // Print calculated N
                    formattedWs.Cells[(i * iOffset) + numExperiments + 2, startPoint + (locationOfAOffset * j)] = cValN;
                    // Print calculated Sqrt N
                    formattedWs.Cells[(i * iOffset) + numExperiments + 2, (startPoint + 1) + (locationOfAOffset * j)] = cValSqrtN;
                    // Print calculated StDev
                    formattedWs.Cells[(i * iOffset) + numExperiments + 3, startPoint + (locationOfAOffset * j)] = cValStDev;
                    // Print calculated SEM
                    formattedWs.Cells[(i * iOffset) + numExperiments + 4, startPoint + (locationOfAOffset * j)] = cSEM;
                    SEMs.Add(cSEM);
                }
            }
        }

        /// <summary>
        /// Produces the "- FINAL" worksheet with means and SEMs for each variable for each sample (except controls).
        /// </summary>
        /// <param name="expData">The list created by parseData that contains all the data from
        /// the raw worksheet.</param>
        /// <author>Chris Meyers</author>
        private void makeFinalWorksheet(List<List<Dictionary<String, Dictionary<String, Double>>>> experiments) {
            int dataStartRow = 2;
            int dataStartCol = 2;

            // Print means and SEMs to final ws
            int counter = 0;
            for (int row = 0; row < numVariables; row++) {
                for (int col = 0; col < 2 * (numSamples - 1); col+=2) {
                    Double currentMean = (Double)means[counter];
                    Double currentSEM = (Double)SEMs[counter];
                    Excel.Range currentMeanLocation = finalWs.Cells[row + dataStartRow, dataStartCol + col];
                    Excel.Range currentSEMLocation = finalWs.Cells[row + dataStartRow, dataStartCol + col + 1];

                    // Checks to see if current mean is a div-by-0 (stored as NaN, printed as 65535)
                    if (Double.IsNaN(currentMean)) {
                        ((Excel.Range)currentMeanLocation).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        ((Excel.Range)currentMeanLocation).Value = 0.001;
                    }
                    else {
                        ((Excel.Range)currentMeanLocation).Value = currentMean;
                    }

                    // Checks to see if current SEM is a div-by-0 (stored as NaN, printed as 65535)
                    if (Double.IsNaN(currentSEM)) {
                        ((Excel.Range)currentSEMLocation).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        ((Excel.Range)currentSEMLocation).Value = 0.001;
                    }
                    else {
                        ((Excel.Range)currentSEMLocation).Value = currentSEM;
                    }
                    
                    counter++;
                }
            }

            // Print variables to final ws on rows
            for (int varName = 0; varName < variableNames.Count; varName++) {
                finalWs.Cells[dataStartRow + varName, 1] = variableNames[varName];
            }

            counter = 0;
            // Print sample name and SEM on columns
            // Ignore first numExperiment many to skip controls.
            for (int sampName = numExperiments; sampName < sampleNames.Count; sampName++) { 
                finalWs.Cells[1, dataStartCol + counter] = sampleNames[sampName];
                finalWs.Cells[1, dataStartCol + counter + 1] = "SEM";
                counter+=2;
            }
        }
    }
}
