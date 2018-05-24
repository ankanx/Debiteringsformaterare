using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Win32;
using System.IO;
using System.Reflection;

namespace Debiteringsformaterare
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    

    public partial class MainWindow : Window
    {
        public static string[] Lines_Global = { };
        public static List<FaktureringsObjekt> FaktureringsObjekt_Global = new List<FaktureringsObjekt>();

        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void Create_Excel_Workfile(object sender, RoutedEventArgs e)
        {
            string file_name = "Debiteringar_" + DateTime.Now.ToString("MMMM") + ".xlsx";
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            var assembly = Assembly.GetExecutingAssembly();
            var resource = "Debiteringsformaterare.Form_appartment.xlsx";

            string out_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + file_name;

            CopyResource(resource, out_path);

            var xlWorkBook = xlApp.Workbooks.Open(out_path);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Console.WriteLine(xlWorkSheet.Rows);

            // Build output
            //for()
            xlWorkSheet.Cells[8, 1] = "ID";
            xlWorkSheet.Cells[9, 2] = "Name";
            xlWorkSheet.Cells[10, 1] = "1";
            xlWorkSheet.Cells[8, 2] = "One";
            xlWorkSheet.Cells[9, 1] = "2";
            xlWorkSheet.Cells[10, 2] = "Two";

            
            try
            {
                xlWorkBook.Save();
                //xlWorkBook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + file_name, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine("No file selected");
                Console.WriteLine(ex);
            }
            xlWorkBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel fil skapad , Du kan hitta den på skrivbordet vid namnet: " + file_name);
        
    }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("worked");
        }

        private void CopyResource(string resourceName, string file)
        {
            using (Stream resource = GetType().Assembly
                                              .GetManifestResourceStream(resourceName))
            {
                if (resource == null)
                {
                    throw new ArgumentException("No such resource", "resourceName");
                }
                using (Stream output = File.OpenWrite(file))
                {
                    resource.CopyTo(output);
                }
            }
        }

        private void Select_File(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            List<FaktureringsObjekt> tmp_list = new List<FaktureringsObjekt>();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.DefaultExt = ".txt";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ShowDialog();

            if ( openFileDialog1.FileName != null )
            {
                try
                {
                    
                    Console.WriteLine(openFileDialog1.FileName);

                    Lines_Global = File.ReadAllLines(openFileDialog1.FileName);

                    foreach (string Line in Lines_Global)
                    {
                        int Nuvarande_Id = 0;
                        FaktureringsObjekt Nuvarande_Objekt = new FaktureringsObjekt();
                        //Console.WriteLine(Line);
                        string[] Delar = Line.Split('\t');
                        int Column = 0;
                        foreach(string str in Delar)
                        {
                            if(str != "")
                            {
                                // Check Id if new object
                                if (Nuvarande_Objekt.Id == 0)
                                {
                                    int.TryParse(str, out Nuvarande_Id);
                                    if (Nuvarande_Id == (int)FaktureringsID.DebiteringsFilensSkapningsDatum)
                                    {
                                        Nuvarande_Objekt.Id = FaktureringsID.DebiteringsFilensSkapningsDatum;
                                        Console.WriteLine("Hittade fakturerings header");

                                    }
                                    else if (Nuvarande_Id == (int)FaktureringsID.BokningsObjekt)
                                    {
                                        Nuvarande_Objekt.Id = FaktureringsID.BokningsObjekt;
                                        Console.WriteLine("Hittade fakturerings boknings objekt");

                                    }
                                    else if (Nuvarande_Id == (int)FaktureringsID.Summa)
                                    {
                                        Nuvarande_Objekt.Id = FaktureringsID.Summa;
                                        Console.WriteLine("Hittade fakturerings summa");

                                    }
                                    else
                                    {
                                        Console.WriteLine("Invalid ID");
                                    }
                                    Column++;
                                    break;
                                }
                  
                                // Else check content
                                switch (Nuvarande_Objekt.Id)
                                {
                                    case FaktureringsID.DebiteringsFilensSkapningsDatum:
                                        if(Column == 1)
                                        {
                                            Nuvarande_Objekt.Fran_Datum = str;
                                        }
                                        if(Column == 2)
                                        {
                                            Nuvarande_Objekt.Till_Datum = str;
                                        }
                                        break;
                                    case FaktureringsID.BokningsObjekt:
                                        if (Column == 1)
                                        {
                                            Nuvarande_Objekt.Lagenhet = str;
                                        }
                                        if (Column == 2)
                                        {
                                            Nuvarande_Objekt.Datum = str;
                                        }
                                        if (Column == 3)
                                        {
                                            Nuvarande_Objekt.Namn = str;
                                        }
                                        if (Column == 4)
                                        {
                                            Nuvarande_Objekt.Typ = str;
                                        }
                                        if (Column == 5)
                                        {
                                            Nuvarande_Objekt.Bokning = str;
                                        }
                                        if (Column == 6)
                                        {
                                            Nuvarande_Objekt.Kostnad = float.Parse(str);
                                        }
                                        break;
                                    case FaktureringsID.Summa:
                                        if (Column == 1)
                                        {
                                            Nuvarande_Objekt.Lagenhet = str;
                                        }
                                        if (Column == 2)
                                        {
                                            Nuvarande_Objekt.Kostnad = float.Parse(str);
                                        }
                                        break;
                                }

                                Column++;


                                Console.WriteLine("|" + str + "|");

                            }
                            
                        }
                        tmp_list.Add(Nuvarande_Objekt);
                    }

                    Console.WriteLine(tmp_list.Count);
                    // Overwrite old object list
                    FaktureringsObjekt_Global = tmp_list;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }

            }
        }

    }
}
