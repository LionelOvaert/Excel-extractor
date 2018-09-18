using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_extractor {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        private FolderBrowserDialog Fbd;
        private string input_folder;
        private string output_folder;
        private ArrayList FichierATraiter = new ArrayList();
        private Excel.Application app;
        private Excel.Workbook wb;
        private Excel.Worksheet downMLT;
        private List<string> headers = new List<string> { "Project Number", "Name Project", "RP", "Phase", "Departement", "Resp. tâche", "Month/Year", "Hours", "Honoraires", "Cout final", "Cout actuel", "Fae", "Facturation", "Commentaire" };
        private List<List<string>> data = new List<List<string>>();
        private int begin_index = 1;
        private int total_entries = 0;
        private string selectedFolderType = "RFP";

        private void MainWindow_Load(object sender, EventArgs e) {

        }

        private void Selection_Click(object sender, RoutedEventArgs e) {
            Fbd = new FolderBrowserDialog();
            if (Fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                input.Clear();
                input_folder = Fbd.SelectedPath.ToString();
                input.Text = input_folder;
            }
        }

        private void Output_Folder_Click(object sender, RoutedEventArgs e) {
            Fbd = new FolderBrowserDialog();
            if (Fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                output.Clear();
                output_folder = Fbd.SelectedPath.ToString();
                output.Text = output_folder;
            }
        }

        private void Conversion_Click(object sender, RoutedEventArgs e) {
            //this.convert_progress.Visibility = Visibility.Visible;
            if (input_folder == null) {
                return;
            }

            this.convert_percentage.Visibility = Visibility.Visible;
            var type = this.folderType.SelectedItem as string;

            switch (type) {
                case "RFP":
                    TraiterRFP(input_folder);
                    break;
                case "Métier":
                    TraiterMetier(input_folder);
                    break;
                case "Responsable":
                    TraiterResponsable(input_folder);
                    break;
                case "Projet":
                    TraiterProjet(input_folder);
                    Debug.WriteLine(FichierATraiter.Count);
                    break;
            }

            //this.convert_progress.Minimum = 1;
            //this.convert_progress.Maximum = this.FichierATraiter.Count;

            app = new Excel.Application();
            wb = app.Workbooks.Add(Missing.Value);
            downMLT = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            try {
                //BackgroundWorker worker = new BackgroundWorker();
                //worker.WorkerReportsProgress = true;
                //worker.DoWork += TraiterFichiers;
                //worker.ProgressChanged += worker_ProgressChanged;

                //worker.RunWorkerAsync();
                TraiterFichiers();
            } catch (Exception ex) {
                Debug.WriteLine(ex);
            } finally {
                wb.Close(true, Missing.Value, Missing.Value);
                app.Quit();
                Marshal.ReleaseComObject(downMLT);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);
                Environment.Exit(0);
            }

        }

        private void TraiterFichiers(/*object sender, DoWorkEventArgs e*/) {

            object objOpt = Missing.Value;
            int fichierCourant = 0;

            foreach (string fichier in FichierATraiter) {
                fichierCourant += 1;
                this.convert_percentage.Text = fichierCourant + "/" + this.FichierATraiter.Count;
                //(sender as BackgroundWorker).ReportProgress(fichierCourant);

                Debug.WriteLine(fichier);
                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                Excel._Worksheet xlWorksheetRecap = null;
                try {
                    xlApp = new Excel.Application();

                    // A supprimer normalement car on flagera le fichier comme une erreur!
                    string ext = Path.GetExtension(fichier);
                    if (Path.GetExtension(fichier) == ".XLSX") {
                        var fichierMod = Path.ChangeExtension(fichier, ".xlsx");
                        File.Move(fichier, fichierMod);
                        xlWorkbook = xlApp.Workbooks.Open(fichierMod, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
                    } else if (ext == ".xlsx") {
                        xlWorkbook = xlApp.Workbooks.Open(fichier, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
                    }
                    // fin suppr

                    if (xlWorkbook == null) {
                        continue;
                    }
                    bool forecastFound = false;
                    bool recapFound = false;
                    foreach (Excel.Worksheet sheet in xlWorkbook.Sheets) {
                        if (sheet.Name.Equals("Forecast")) {
                            forecastFound = true;
                        }
                        if (sheet.Name.Equals("Forecast")) {
                            recapFound = true;
                        }
                    }
                    if (forecastFound && recapFound) {
                        xlWorksheet = xlWorkbook.Worksheets["Forecast"];
                        xlWorksheetRecap = xlWorkbook.Worksheets["Récap financier"];
                    } else {
                        continue;
                    }

                    int nbRow = xlWorksheet.Cells[xlWorksheet.Rows.Count, 2].End(Excel.XlDirection.xlUp).Row;

                    int row1 = 9;
                    string cell1 = "A" + row1.ToString();
                    string cell2 = "AQ" + nbRow.ToString();
                    Excel.Range rng = xlWorksheet.get_Range(cell1, cell2);
                    object[,] cells = (object[,])rng.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                    row1 = 2;
                    cell1 = "D" + row1.ToString();
                    cell2 = "K3";
                    rng = xlWorksheetRecap.get_Range(cell1, cell2);
                    object[,] cellsInfo = (object[,])rng.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                    row1 = 14;
                    cell1 = "A" + row1.ToString();
                    cell2 = "N" + row1.ToString();
                    rng = xlWorksheetRecap.get_Range(cell1, cell2);
                    object[,] cellsRecap = (object[,])rng.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                    xlWorkbook.Close(false, objOpt, objOpt);
                    xlApp.Quit();

                    int nbCol = 43;
                    string projectNumber = cellsInfo[1, 1] + "";
                    string projectName = cellsInfo[1, 8] + "";
                    string metier = cellsInfo[2, 1] + "";
                    string RP = cellsInfo[2, 8] + "";
                    string resp = "";
                    string phase = "";
                    string mois_annee = "";
                    string heures = "";
                    double cout_restime = Convert.ToDouble(cellsRecap[1, 4]);
                    Debug.WriteLine("cout restime" + cout_restime);
                    double cout_actuel = Convert.ToDouble(cellsRecap[1, 5]);
                    double revenus_estimes_total = Convert.ToDouble(cellsRecap[1, 6]);
                    double fae = Convert.ToDouble(cellsRecap[1, 13]);
                    double facturation = Convert.ToDouble(cellsRecap[1, 7]);
                    string commentaires = cellsRecap[1, 14] + "";

                    int yearCol = 1;
                    for (int i = yearCol + 1; i <= nbRow - 9; i++) {
                        // TODO: Vérifier quelles opérations prennent du temps => optimisation nécessaire car procesus trop long (pour 3 fichiers seulement)
                        // UPDATE: Une lecture au début de la zone intéressante du fichier améliore les perfs un peu 
                        //Debug.WriteLine(i+", "+cells[i, 2]);
                        if (cells[i, 2] != null) {
                            if (!"Code".Equals(cells[i, 2].ToString())) {
                                string dep = cells[i, 6] + "";
                                if ("".Equals(dep) || "S Hours".Equals(dep) || dep.Contains("h")) {
                                    continue;
                                }
                                if ((int.Parse(cells[i, 6].ToString())) > 0) {
                                    metier = cells[i, 2] + "";
                                    resp = cells[i, 3] + "";

                                    for (int j = 7; j <= nbCol; j++) {
                                        if (cells[i, j] != null && (int.Parse(cells[i, j].ToString())) > 0) {
                                            mois_annee = ((DateTime)cells[yearCol, j]).ToString("MM/dd/yyyy");
                                            heures = ((double)cells[i, j]).ToString();
                                            data.Add(new List<string> { projectNumber, projectName, RP, phase, metier, resp, mois_annee, heures });
                                            total_entries++;
                                        }
                                    }
                                }
                            } else {
                                phase = cells[i + 1, 1] + "";
                            }
                        }
                    }
                    Debug.WriteLine(revenus_estimes_total);
                    Debug.WriteLine((double)total_entries);
                    Debug.WriteLine(Math.Round(revenus_estimes_total / (double)total_entries, 2, MidpointRounding.AwayFromZero));
                    Debug.WriteLine(Math.Round(revenus_estimes_total / (double)total_entries, 2, MidpointRounding.AwayFromZero) + "");
                    Debug.WriteLine("-------------------------");
                    string honoraires = Math.Round(revenus_estimes_total / (double)total_entries, 2, MidpointRounding.AwayFromZero) + "";
                    string cout_final = Math.Round(cout_restime / (double)total_entries, 2, MidpointRounding.AwayFromZero) + "";
                    string cout_actuel_str = Math.Round(cout_actuel / (double)total_entries, 2, MidpointRounding.AwayFromZero) + "";
                    string fae_str = Math.Round(fae / (double)total_entries, 2, MidpointRounding.AwayFromZero) + "";
                    string facturation_str = Math.Round(facturation / (double)total_entries, 2, MidpointRounding.AwayFromZero) + "";
                    foreach (List<string> row in data) {
                        row.Add(honoraires);
                        row.Add(cout_final);
                        row.Add(cout_actuel_str);
                        row.Add(fae_str);
                        row.Add(facturation_str);
                        row.Add(commentaires);
                    }

                } catch (Exception ex) {
                    Debug.WriteLine(ex);
                    throw new Exception(ex.Message);
                } finally {
                    if (xlWorksheet != null) {
                        Marshal.ReleaseComObject(xlWorksheet);
                    }
                    if (xlWorkbook != null) {
                        Marshal.ReleaseComObject(xlWorkbook);
                    }
                    if (xlApp != null) {
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
            }

            // Effectuer toutes les écritures à la fin
            object[,] arr = new object[1 + data.Count, headers.Count];

            //HEADERS
            for (int i = 0; i < headers.Count; i++) {
                arr[0, i] = headers[i];
            }

            //foreach (List<string> row in data) {
            //    foreach(string item in row) {
            //        Debug.WriteLine(item);
            //    }
            //}

            //CONTENU
            int counter = 0;
            foreach (List<string> row in data) {
                for (int i = 0; i < row.Count; i++) {
                    //Debug.WriteLine(row.Count+","+begin_index+counter+","+i);
                    if (i >= 8 && i <= 12) {
                        arr[begin_index + counter, i] = double.Parse(row[i]);
                    } else {
                        arr[begin_index + counter, i] = row[i];
                    }
                }
                //arr[begin_index + counter, row.Length-1] = Double.Parse(row[row.Length - 1]);
                counter += 1;
            }

            //Excel.Range rg = downMLT.Cells[2, 7];
            //rg.EntireColumn.NumberFormat = "mmm-yy";
            downMLT.Range["G2","G"+counter+1].NumberFormat = "mmm-yy";
            downMLT.Range["I2","I"+counter+1].NumberFormat = "# ##0.00 €";
            downMLT.Range["J2", "J" + counter+1].NumberFormat = "# ##0.00 €";
            downMLT.Range["K2", "K" + counter+1].NumberFormat = "# ##0.00 €";
            downMLT.Range["L2", "L" + counter+1].NumberFormat = "# ##0.00 €";
            downMLT.Range["M2", "M" + counter+1].NumberFormat = "# ##0.00 €";
            // Nécessaire de nommer ?
            downMLT.Name = "down MLT";

            Excel.Range c1 = downMLT.Cells[1, 1];
            Excel.Range c2 = downMLT.Cells[1 + data.Count, headers.Count];
            Excel.Range range = downMLT.get_Range(c1, c2);
            range.Value = arr;

            for (int i = 1; i <= headers.Count; i++) {
                downMLT.Columns[i].AutoFit();
            }


            var final_output_name = this.output_name.Text;
            if (this.output_name != null && this.output_name.Equals("")) {
                final_output_name = "MLT";
            }

            wb.SaveAs(output.Text + "\\" + final_output_name + ".xlsm", Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlNoChange, objOpt, objOpt, objOpt, objOpt, objOpt);
            System.Windows.MessageBox.Show("Conversion terminée");
        }

        private void TraiterRFP(string rfp) {
            string[] metiers = Directory.GetDirectories(rfp);
            foreach (string metier in metiers) {
                if (metier != null && !metier.Contains("ARCHIVES") && !metier.Equals("DATA")) {
                    TraiterMetier(metier);
                }
            }
        }

        private void TraiterMetier(string metier) {
            string[] responsables = Directory.GetDirectories(metier);
            foreach (string responsable in responsables) {
                TraiterResponsable(responsable);
            }
        }

        private void TraiterResponsable(string responsable) {
            string[] projets = Directory.GetDirectories(responsable);
            foreach (string projet in projets) {
                TraiterProjet(projet);
            }
        }

        private void TraiterProjet(string projet) {
            //Obtenir des fichiers au lieu de string
            //Mettre les fichiers dans une list
            //Trier et prendre le premier
            var dir = new DirectoryInfo(projet);
            FileInfo[] files = dir.GetFiles();
            if (files.Length == 0) {
                return;
            }
            if (files.Length == 1) {
                FichierATraiter.Add(files[0].FullName);
                return;
            }
            DateTime[] dates = new DateTime[files.Length];
            for (int i = 0; i < files.Length; i++) {
                if (files[i].LastWriteTime.Year <= 1601) {
                    dates[i] = files[i].CreationTime;
                } else {
                    dates[i] = files[i].LastWriteTime;
                }
            }
            Array.Sort(dates, files);
            FichierATraiter.Add(files.Last().FullName);
        }

        private void folderType_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) {
            var comboBox = sender as System.Windows.Controls.ComboBox;

            string value = comboBox.SelectedItem as string;
            this.selectedFolderType = value;
            this.Title = "Selected: " + value;
        }

        private void folderType_Loaded(object sender, RoutedEventArgs e) {
            List<string> data = new List<string>();
            data.Add("RFP");
            data.Add("Métier");
            data.Add("Responsable");
            data.Add("Projet");

            var folderType = sender as System.Windows.Controls.ComboBox;
            folderType.ItemsSource = data;
            folderType.SelectedIndex = 0;
        }

        //void worker_ProgressChanged(object sender, ProgressChangedEventArgs e) {
        //    convert_progress.Value = e.ProgressPercentage;
        //}
    }
}
