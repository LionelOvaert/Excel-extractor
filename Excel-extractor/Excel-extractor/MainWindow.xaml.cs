using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
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

            int k = 2;
            object objOpt = Missing.Value;

            downMLT.Cells[1, 1] = "Project Number";
            downMLT.Cells[1, 2] = "Name Project";
            downMLT.Cells[1, 3] = "RP";
            downMLT.Cells[1, 4] = "Phase";
            downMLT.Cells[1, 5] = "Departement";
            downMLT.Cells[1, 6] = "Resp. tâche";
            downMLT.Cells[1, 7] = "Month/Year";
            downMLT.Cells[1, 8] = "Hours";

            Excel.Range rg = downMLT.Cells[2, 7];
            rg.EntireColumn.NumberFormat = "mmm-yy";

            foreach (string fichier in FichierATraiter) {

                Debug.WriteLine(fichier);

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = null;

                string ext = Path.GetExtension(fichier);
                if (Path.GetExtension(fichier) == ".XLSX") {
                    var fichierMod = Path.ChangeExtension(fichier, ".xlsx");
                    File.Move(fichier, fichierMod);
                    xlWorkbook = xlApp.Workbooks.Open(fichierMod, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
                } else if(ext == ".xlsx"){
                    xlWorkbook = xlApp.Workbooks.Open(fichier, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
                }

                if(xlWorkbook == null) {
                    continue;
                }

                Excel._Worksheet xlWorksheet = xlWorkbook.Worksheets["Forecast"];
                try {
                    int yearCol = 0;

                    for(int i = 1; i < 20; i++) {
                        if(xlWorksheet.Cells[i, 7].Value2 != null && xlWorksheet.Cells[i,7].Value2.ToString() != "") {
                            yearCol = i;
                            break;
                        }
                    }

                    int nbCol = 6 + (int)xlApp.WorksheetFunction.CountA(xlWorksheet.get_Range((Excel.Range)xlWorksheet.Cells[yearCol, 7], (Excel.Range)xlWorksheet.Cells[yearCol, 95]));
                    int nbRow = (int)xlApp.WorksheetFunction.CountA(xlWorksheet.get_Range((Excel.Range)xlWorksheet.Cells[yearCol+5, 5], (Excel.Range)xlWorksheet.Cells[1000, 5]));
                    string projectNumber = "";
                    string projectName = "";
                    string metier = "";
                    string resp = "";
                    string mois_annee_str = "";
                    double heures = 0;

                    projectNumber = xlWorksheet.Cells[1, 2].Value2 + "";
                    projectName = xlWorksheet.Cells[2, 2].Value2 + "";

                    //nbRow--;
                    for (int i = yearCol+5; i <= nbRow; i++) {
                        string dep = xlWorksheet.Cells[i, 6].Value2 + "";
                        if(dep == "" || dep == "S Hours" || dep.Contains("h")) {
                            continue;
                        }
                        if ((Int32.Parse(xlWorksheet.Cells[i, 6].Value2.ToString())) > 0) {
                            metier = xlWorksheet.Cells[i, 2].Value2 + "";
                            resp = xlWorksheet.Cells[i, 3].Value2 + "";

                            for (int j = 7; j < nbCol; j++) {
                                if (xlWorksheet.Cells[i, j].Value2 > 0) {
                                    mois_annee_str = xlWorksheet.Cells[yearCol+2, j].Value2.ToString();
                                    double date = double.Parse(mois_annee_str);
                                    var mois_annee = DateTime.FromOADate(date).ToString("dd/MM/yyyy");
                                    heures = xlWorksheet.Cells[i, j].Value2;
                                    Debug.WriteLine(heures);
                                    downMLT.Cells[k, 1] = projectNumber;
                                    downMLT.Cells[k, 2] = projectName;
                                    //downMLT.Cells[k, 3] = RP;
                                    downMLT.Cells[k, 3] = "";
                                    downMLT.Cells[k, 5] = metier;
                                    downMLT.Cells[k, 6] = resp;
                                    downMLT.Cells[k, 7] = mois_annee;
                                    downMLT.Cells[k, 8] = heures;
                                    k++;
                                }
                            }
                        }
                    }
                } catch (Exception ex) {
                    Debug.WriteLine(ex);
                    throw new Exception(ex.Message);
                } finally {
                    xlWorkbook.Close(false, objOpt, objOpt);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlWorkbook);
                    Marshal.ReleaseComObject(xlApp);
                }
            }
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
