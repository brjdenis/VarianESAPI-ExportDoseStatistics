using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using Microsoft.Win32;

//[assembly: AssemblyVersion("0.0.0.1")]


namespace VMS.TPS
{
        public class Script
        {
        // Taken from https://medium.com/@jhmcastelo/2-steps-to-plot-your-esapi-script-data-80d3e56e7cd
        // var createHTML = HTMLBuilder.StaticText(x, y, z);
        // var runner = new HTMLRunner(createHTML);
        // runner.Launch("Test");
        public class HTMLRunner
        {
            public string Text { get; set; }
            public string TempFolder { get; set; }
            public HTMLRunner(string text)
            {
                TempFolder = Path.GetTempPath();
                Text = text.ToString();
            }
            public void Launch(string title)
            {
                var fileName = Path.Combine(TempFolder, title + ".html");
                File.WriteAllText(fileName, Text);
                System.Diagnostics.Process.Start(fileName);
            }
        }

        public class HTMLBuilder //Generate the text for the plot to be passed to HTMLRunner class
        {
            public static string StaticText(List<List<string>> table, List<List<string>> table2, List<List<string>> table3,
                    List<List<string>> table4)
            {
                var preX = @"
                <!DOCTYPE html>
                <html lang='en'>
                <head>
                    <meta charset=utf-8>
                    <style>
                    table, th, td {
                      border: 1px solid black;
                      border-collapse: collapse;
                    }
                    th, td {
                      padding: 7px;
                    }
                    </style>

                     </head>
                <body>
                ";

                var posY = @"
                </body>
                </html>
                ";

                // Plan objectives left in Eclipse
                string tablehtml = "<h1>Prescription</h1><table>";

                foreach (var column in table)
                {
                    tablehtml += "<tr>";

                    foreach (var row in column)
                    {
                        tablehtml += "<td>" + row + "</td>";
                    }

                    tablehtml += "</tr>";
                }
                tablehtml += "</table>";

                // Plan objectives right in Eclipse
                string tablehtml2 = "<h1>Measure</h1><table>";

                foreach (var column in table2)
                {
                    tablehtml2 += "<tr>";

                    foreach (var row in column)
                    {
                        tablehtml2 += "<td>" + row + "</td>";
                    }

                    tablehtml2 += "</tr>";
                }
                tablehtml2 += "</table>";

                // DVH 1
                string tablehtml3 = "<table>";
                foreach (var column in table3)
                {
                    tablehtml3 += "<tr>";
                    foreach (var row in column)
                    {
                        tablehtml3 += "<td>" + row + "</td>";
                    }
                    tablehtml3 += "</tr>";
                }
                tablehtml3 += "</table>";

                // DVH 2
                string tablehtml4 = "<table>";
                foreach (var column in table4)
                {
                    tablehtml4 += "<tr>";
                    foreach (var row in column)
                    {
                        tablehtml4 += "<td>" + row + "</td>";
                    }
                    tablehtml4 += "</tr>";
                }
                tablehtml4 += "</table>";

                // Dropdown menu
                string dropdown =
                @"  <h1>DVH</h1><form>
                Choose presentation:
                <select id = 'selected' onchange='changeselect();'>
                <option id = '1' selected> cGy </option>
                <option id = '2'> % </option>
                </select>
                </form>
                <br>
                <div id='table'>
                </div>
                <script>
                    var table1 = '" + tablehtml3 + @"';
                    var table2 = '" + tablehtml4 + @"';
                    changeselect();
                    function changeselect(){
                        var ind = document.getElementById('selected').selectedIndex+1;
                        if (ind==1){
                            document.getElementById('table').innerHTML = table1;
                        }
                        if (ind==2){
                            document.getElementById('table').innerHTML = table2;
                        }
                    }
                </script>
                ";
                return preX + tablehtml + tablehtml2 + dropdown + posY;
            }
        }

        public Script()
        {
        }

        private List<List<string>> BuildMessage(DoseValuePresentation dosevaluepresentation, VolumePresentation volumepresentation,
            ExternalPlanSetup plan, string filepath)
        {
            // DVH stats 
            string doseunit = "cGy";

            // Read stats from table
            List<string> target = new List<string>();
            List<string> source = new List<string>();
            List<string> targetunit = new List<string>();
            List<string> sourceunit = new List<string>();
            
            using (var reader = new StreamReader(filepath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';').Select(p => p.Trim()).ToList();

                    target.Add(values[0]);
                    source.Add(values[1]);
                    targetunit.Add(values[2]);
                    sourceunit.Add(values[3]);
                }
            }

            List<List<string>> message = new List<List<string>> { };

            List<string> listofindexes = new List<string> {
                "StructureID",
                "Volume [cm3]",
                "Dose cover [%]",
                "Sampling cover [%]",
                String.Format("Min Dose [{0}]", doseunit),
                String.Format("Max Dose [{0}]", doseunit),
                String.Format("Mean Dose [{0}]", doseunit) 
            };

            for ( int i = 0; i < target.Count(); i++)
            {
                var target2 = target[i];
                var targetunit2 = targetunit[i];
                var source2 = source[i];
                var sourceunit2 = sourceunit[i];
                string s = "";

                if(target[i] == "V")
                {
                    if (targetunit2 == "abs")
                    {
                        if (sourceunit2 == "abs")
                        {
                            s = String.Format("V{0}{1} [{2}]", source2, "cGy", "cm3");
                        }
                        else
                        {
                            s = String.Format("V{0}{1} [{2}]", source2, "%", "cm3");
                        }
                    }
                    else
                    {
                        if (sourceunit2 == "abs")
                        {
                            s = String.Format("V{0}{1} [{2}]", source2, "cGy", "%");
                        }
                        else
                        {
                            s = String.Format("V{0}{1} [{2}]", source2, "%", "%");
                        }
                    }
                }
                else
                {
                    if (targetunit2 == "abs")
                    {
                        if (sourceunit2 == "abs")
                        {
                            s = String.Format("D{0}{1} [{2}]", source2, "cm3", "cGy");
                        }
                        else
                        {
                            s = String.Format("D{0}{1} [{2}]", source2, "%", "cGy");
                        }
                    }
                    else
                    {
                        if (sourceunit2 == "abs")
                        {
                            s = String.Format("D{0}{1} [{2}]", source2, "cm3", "%");
                        }
                        else
                        {
                            s = String.Format("D{0}{1} [{2}]", source2, "%", "%");
                        }
                    }
                }
                listofindexes.Add(s);
            }
            message.Add(listofindexes);

            foreach (var s in plan.StructuresSelectedForDvh.OrderBy(x => x.Id))
            {
                DVHData dvh = plan.GetDVHCumulativeData(s, dosevaluepresentation, volumepresentation, 0.1);
                string id = s.Id;
                string volume = string.Format("{0:0.0}", s.Volume);
                string dosecover = string.Format("{0:0.0}", dvh.Coverage * 100.0);
                string samplingcover = string.Format("{0:0.0}", dvh.SamplingCoverage * 100.0);
                string mindose = string.Format("{0:0.0}", dvh.MinDose.Dose);
                string maxdose = string.Format("{0:0.0}", dvh.MaxDose.Dose);
                string meandose = string.Format("{0:0.0}", dvh.MeanDose.Dose);

                List<string> results = new List<string>
                {
                    id, volume, dosecover, samplingcover, mindose, maxdose, meandose
                };

                for (int i = 0; i < target.Count(); i++)
                {
                    var target2 = target[i];
                    var targetunit2 = targetunit[i];
                    var source2 = source[i];
                    var sourceunit2 = sourceunit[i];
                    string r = "";

                    if (target[i] == "V")
                    {
                        if (targetunit2 == "abs")
                        {
                            if (sourceunit2 == "abs")
                            {
                                r = string.Format("{0:0.00}", plan.GetVolumeAtDose(s, new DoseValue(Convert.ToDouble(source2), DoseValue.DoseUnit.cGy), VolumePresentation.AbsoluteCm3));
                            }
                            else
                            {
                                r = string.Format("{0:0.00}", plan.GetVolumeAtDose(s, new DoseValue(Convert.ToDouble(source2), DoseValue.DoseUnit.Percent), VolumePresentation.AbsoluteCm3));
                            }
                        }
                        else
                        {
                            if (sourceunit2 == "abs")
                            {
                                r = string.Format("{0:0.00}", plan.GetVolumeAtDose(s, new DoseValue(Convert.ToDouble(source2), DoseValue.DoseUnit.cGy), VolumePresentation.Relative));
                            }
                            else
                            {
                                r = string.Format("{0:0.00}", plan.GetVolumeAtDose(s, new DoseValue(Convert.ToDouble(source2), DoseValue.DoseUnit.Percent), VolumePresentation.Relative));
                            }
                        }
                    }
                    else
                    {
                        if (targetunit2 == "abs")
                        {
                            if (sourceunit2 == "abs")
                            {
                                r = string.Format("{0:0.00}", plan.GetDoseAtVolume(s, Convert.ToDouble(source2), VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose);
                            }
                            else
                            {
                                r = string.Format("{0:0.00}", plan.GetDoseAtVolume(s, Convert.ToDouble(source2), VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose);
                            }
                        }
                        else
                        {
                            if (sourceunit2 == "abs")
                            {
                                r = string.Format("{0:0.00}", plan.GetDoseAtVolume(s, Convert.ToDouble(source2), VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative).Dose);
                            }
                            else
                            {
                                r = string.Format("{0:0.00}", plan.GetDoseAtVolume(s, Convert.ToDouble(source2), VolumePresentation.Relative, DoseValuePresentation.Relative).Dose);
                            }
                        }
                    }
                    results.Add(r);
                }
                message.Add(results);
            }
            return message;
        }


        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            ExternalPlanSetup plan = context.ExternalPlanSetup;

            List<List<string>> message = new List<List<string>> { };
            List<List<string>> message2 = new List<List<string>> { };

            string filepath = "";

            if (filepath == "")
            {
                OpenFileDialog choofdlog = new OpenFileDialog();
                choofdlog.Filter = "All Files (*.*)|*.*";
                choofdlog.FilterIndex = 1;
                choofdlog.Multiselect = false;

                if ((bool)choofdlog.ShowDialog())
                {
                    filepath = choofdlog.FileName;          
                }
                else
                {
                    return;
                }
            }

            if (plan.ProtocolID != "")
            {
                List<ProtocolPhasePrescription> protocolprescription = new List<ProtocolPhasePrescription> { };
                List<ProtocolPhaseMeasure> protocolmeasure = new List<ProtocolPhaseMeasure> { };
                plan.GetProtocolPrescriptionsAndMeasures(ref protocolprescription, ref protocolmeasure);

                // First prescription
                var firstelement = protocolprescription.First();
                message.Add(new List<string> { 
                    "StructureID", 
                    "PrescModifier", 
                    "PrescParam",
                    "TargetFractionDose [" + firstelement.TargetFractionDose.UnitAsString + "]",
                    "TargetTotalDose [" + firstelement.TargetTotalDose.UnitAsString + "]",
                    "ActualTotalDose ["+  firstelement.ActualTotalDose.UnitAsString + "]",
                    "TargetIsMet?"
                });

                foreach (var p in protocolprescription)
                {
                    string id = p.StructureId;
                    string prescmodifier = p.PrescModifier.ToString().Replace("PrescriptionModifier", "");
                    string prescparam = p.PrescParameter.ToString();
                    string targfracdose = string.Format("{0:0.00}", p.TargetFractionDose.Dose);
                    string targtotdose = string.Format("{0:0.00}", p.TargetTotalDose.Dose);
                    string acttotdose = string.Format("{0:0.00}", p.ActualTotalDose.Dose);
                    string status = p.TargetIsMet.ToString();

                    message.Add(new List<string> { id, prescmodifier, prescparam, targfracdose, targtotdose, acttotdose, status });
                }

                //New table
                message2.Add(new List<string> { "StructureID", "Param", "Modifier", "TargetValue", "ActualValue", "TargetIsMet" });

                foreach (var p in protocolmeasure)
                {
                    string id = p.StructureId;
                    string modifier = p.Modifier.ToString().Replace("MeasureModifier", "");
                    string param = p.TypeText;
                    string targetvalue = p.TargetValue.ToString();
                    string actualvalue = string.Format("{0:0.00}", p.ActualValue);
                    string targetismet = p.TargetIsMet.ToString();

                    message2.Add(new List<string> { id, param, modifier, targetvalue, actualvalue, targetismet });
                }
            }

            // DVH stats Gy/cm3
            List<List<string>> message3 = BuildMessage(DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, plan, filepath);

            // DVH stats %/cm3
            List<List<string>> message4 = BuildMessage(DoseValuePresentation.Relative, VolumePresentation.AbsoluteCm3, plan, filepath);

            var createHTML = HTMLBuilder.StaticText(message, message2, message3, message4);
            var runner = new HTMLRunner(createHTML);
            runner.Launch("DoseStats");
        }
    }
}
