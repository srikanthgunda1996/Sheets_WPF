using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using win = System.Windows.Forms;
using Autodesk.Revit.DB.Structure;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
using System.Reflection;

namespace sHEETS_WPF
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class MyForm : Window
    {
        ObservableCollection<Dataclass> dataclasses { get; set; }

        public ObservableCollection<Element> titleblock { get; set; }

        public ObservableCollection<View> view { get; set; }
        public MyForm(List<Element> titleblocks, List<View> views)
        {
            InitializeComponent();


            titleblock = new ObservableCollection<Element>(titleblocks);
            dataclasses = new ObservableCollection<Dataclass>();
            view = new ObservableCollection<View>(views);
            //dataclasses.Add(new Dataclass());

            gridData.ItemsSource = dataclasses;
            paperSize.ItemsSource = titleblock;
            viewList.ItemsSource = view;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.DialogResult=false;this.Close();
        }

        public List<Dataclass> GetData()
        {
            return dataclasses.ToList();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            dataclasses.Add(new Dataclass());
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (Dataclass dataclass in dataclasses)
                {
                    if (gridData.SelectedItem == dataclass)
                    {
                        dataclasses.Remove(dataclass);
                    }
                }
            }
            catch (Exception) { }

        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
            dialog.Title = "Select Excel File";
            dialog.Filter = "Excel files | *.xlsx;*.xls;*.xlsm";
            dialog.Multiselect = false;

            List<List<string>> exceldata = new List<List<string>>();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;

                excel.Application exc = new excel.Application();
                excel.Workbook wrkbook = exc.Workbooks.Open(file);
                excel.Worksheet wrksheet = wrkbook.Worksheets[1];
                excel.Range rnge = wrksheet.UsedRange;

                int row = rnge.Rows.Count;
                int colmn = rnge.Columns.Count;



                for (int i = 1; i <= row; i++)
                {
                    List<string> rowdata = new List<string>();
                    for (int j = 1; j <= colmn; j++)
                    {
                        string cellcontent = wrksheet.Cells[i, j].Value.ToString();
                        rowdata.Add(cellcontent);
                    }
                    exceldata.Add(rowdata);
                }

                   exceldata.RemoveAt(0);
                
            }

            foreach (List<string> curRow in exceldata)
            {
                    Dataclass curData = new Dataclass();
                    curData.Column1 = curRow[0].ToString();
                    curData.Column2 = curRow[1].ToString();
                    if (curRow[2].ToString() == "True")
                    {
                        curData.Column3 = true;
                    }
                    else
                    {
                        curData.Column3 = false;
                    }
                foreach(Element curSheet in titleblock)
                {
                    if (curSheet.Name == curRow[3].ToString())
                    {
                        curData.Column4 = curSheet;
                    }
                }

                foreach (View curView in view)
                {
                    if (curView.Name == curRow[4].ToString())
                    {
                        curData.Column5 = curView;
                    }
                }
                
                dataclasses.Add(curData);
            }

        }



        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            string filepath = "";
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filepath = dialog.SelectedPath + "\\exportedfile.xlsx";


                excel.Application excelApp = new excel.Application();
                excelApp.Visible = true;

                // Create a new workbook and worksheet
                excel.Workbook workbook = excelApp.Workbooks.Add();
                excel.Worksheet worksheet = workbook.Sheets[1];
                int i = 1;
                foreach (Dataclass dataclass in dataclasses)
                {
                        worksheet.Cells[i, 1].Value = dataclass.Column1.ToString();
                        worksheet.Cells[i, 2].Value = dataclass.Column2.ToString();
                        worksheet.Cells[i, 3].Value = dataclass.Column3.ToString();
                        //string test1 = dataclass.Column4.Name;
                        //string test2 = dataclass.Column5.Name;
                        try
                        {
                            if (dataclass.Column4 != null)
                                worksheet.Cells[i, 4].Value = dataclass.Column4.Name;
                        }
                        catch { }
                        try
                        {
                            if (dataclass.Column5 != null)
                            {
                                worksheet.Cells[i, 5].Value = dataclass.Column5.Name;
                            }
                            else { }
                        }
                        catch { }
                   
                        
                        i++;
                    
                }
                workbook.Save();

            }
        }
    }

    public class Dataclass
    {
        public string Column1 { get; set; }
        public string Column2 { get; set; }
        public bool Column3 { get; set; }
        public Element Column4 { get; set; }

        public View Column5 { get; set; }

    }
}
