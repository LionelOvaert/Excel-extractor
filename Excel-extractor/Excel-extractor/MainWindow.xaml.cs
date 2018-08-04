using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
        private string[] headers = { "Project Number", "Name Project", "RP", "Phase", "Departement", "Resp. tâche", "Month/Year", "Hours" };
        private List<string[]> data = new List<string[]>();
        private int begin_index = 1;
        private int total_entries = 0;

        private void MainWindow_Load(object sender, EventArgs e) {
            //if (!CheckDatabaseExist()) {
            //    GenerateDatabase();
            //}
        }

        //private void GenerateDatabase() {
        //    List<string> cmds = new List<string>();
        //    if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Script.sql")) {
        //        TextReader tr = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\Script.sql");
        //        string line = "";
        //        string cmd = "";
        //        while ((line = tr.ReadLine()) != null) {
        //            if (line.Trim().ToUpper() == "GO") {
        //                cmds.Add(cmd);
        //                cmd = "";
        //            } else {
        //                cmd += line + "\r\n";
        //            }
        //        }
        //        if (cmds.Count > 0) {
        //            SqlCommand command = new SqlCommand();
        //            command.Connection = new SqlConnection(@"Data Source=.\sqlexpress;Initial Catalog=MASTER;Integrated Security=True");
        //            command.CommandType = System.Data.CommandType.Text;
        //            command.Connection.Open();
        //            for (int i = 0; i < cmds.Count; i++) {
        //                command.CommandText = cmds[i];
        //                command.ExecuteNonQuery();
        //            }
        //        }
        //    }
        //}

        //private bool CheckDatabaseExist() {
        //    SqlConnection Connection = new SqlConnection(@"Data Source=.\sqlexpress;Initial Catalog=Projets;Integrated Security=True");
        //    try {
        //        Connection.Open();
        //        return true;
        //    } catch {
        //        return false;
        //    }
        //}

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
            if (input_folder == null) {
                return;
            }
            string[] metiers = Directory.GetDirectories(input_folder);
            foreach (string metier in metiers) {
                TraiterMetier(metier);
            }
            app = new Excel.Application();
            wb = app.Workbooks.Add(Missing.Value);
            downMLT = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            try {
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

        private void TraiterFichiers() {

            object objOpt = Missing.Value;

            foreach (string fichier in FichierATraiter) {

                Debug.WriteLine(fichier);
                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                try {
                    xlApp = new Excel.Application();

                    string ext = Path.GetExtension(fichier);
                    if (Path.GetExtension(fichier) == ".XLSX") {
                        var fichierMod = Path.ChangeExtension(fichier, ".xlsx");
                        File.Move(fichier, fichierMod);
                        xlWorkbook = xlApp.Workbooks.Open(fichierMod, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
                    } else if (ext == ".xlsx") {
                        xlWorkbook = xlApp.Workbooks.Open(fichier, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
                    }

                    if (xlWorkbook == null) {
                        continue;
                    }

                    xlWorksheet = xlWorkbook.Worksheets["Forecast"];
                    int yearCol = 0;

                    int nbRow = xlWorksheet.Cells[xlWorksheet.Rows.Count, 2].End(Excel.XlDirection.xlUp).Row;

                    int row1 = 1;
                    string cell1 = "A" + row1.ToString();
                    string cell2 = "AQ" + nbRow.ToString();
                    Excel.Range rng = xlWorksheet.get_Range(cell1, cell2);
                    object[,] cells = (object[,])rng.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                    xlWorkbook.Close(false, objOpt, objOpt);
                    xlApp.Quit();

                    for (int i = 1; i < 20; i++) {
                        if (cells[i, 7] != null && cells[i, 7].ToString() != "") {
                            yearCol = i + 2;
                            break;
                        }
                    }

                    int nbCol = 43;
                    string projectNumber = cells[1, 2] + "";
                    string projectName = cells[2, 2] + "";
                    string metier = "";
                    string resp = "";
                    string mois_annee = "";
                    string heures = "";

                    for (int i = yearCol + 1; i <= nbRow; i++) {
                        // TODO: Vérifier quelles opérations prennent du temps => optimisation nécessaire car procesus trop long (pour 3 fichiers seulement)
                        // UPDATE: Une lecture au début de la zone intéressante du fichier améliore les perfs un peu 

                        if (cells[i, 2] != null && cells[i, 2].ToString() != "Code") {
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
                                        data.Add(new string[] { projectNumber, projectName, "", "", metier, resp, mois_annee, heures });
                                        total_entries++;
                                    }
                                }
                            }
                        }
                    }
                } catch (Exception ex) {
                    Debug.WriteLine(ex);
                    throw new Exception(ex.Message);
                } finally {
                    //if (xlWorkbook != null) {
                    //    xlWorkbook.Close(false, objOpt, objOpt);
                    //}
                    //if (xlApp != null) {
                    //    xlApp.Quit();
                    //}
                    Marshal.ReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlWorkbook);
                    Marshal.ReleaseComObject(xlApp);
                }
            }

            // Effectuer toutes les écritures à la fin
            object[,] arr = new object[1 + data.Count, 8];

            //HEADERS
            for (int i = 0; i < 8; i++) {
                arr[0, i] = headers[i];
            }

            //CONTENU
            int counter = 0;
            foreach (string[] row in data) {
                for (int i = 0; i < row.Length; i++) {
                    arr[begin_index + counter, i] = row[i];
                }
                //arr[begin_index + counter, row.Length-1] = Double.Parse(row[row.Length - 1]);
                counter++;
            }

            Excel.Range c1 = downMLT.Cells[1, 1];
            Excel.Range c2 = downMLT.Cells[1 + data.Count, 8];
            Excel.Range range = downMLT.get_Range(c1, c2);
            range.Value = arr;

            Excel.Range rg = downMLT.Cells[2, 7];
            rg.EntireColumn.NumberFormat = "mmm-yy";
            downMLT.Name = "down MLT";

            wb.SaveAs(output.Text + "\\MLT.xlsm", Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, objOpt, objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlNoChange, objOpt, objOpt, objOpt, objOpt, objOpt);
            System.Windows.MessageBox.Show("Conversion terminée");
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


    }
}
