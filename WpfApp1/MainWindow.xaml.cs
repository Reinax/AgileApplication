using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string Topmeal;
        private string SearchItemOne;
        private string SearchItemTwo;
        private void ButtonExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.InitialDirectory = "b:\\Temp";
            fileDialog.Filter = "Excel Files(.xlsx)|*.xlsx";
            fileDialog.Title = "Select an excel file";
            fileDialog.RestoreDirectory = true;

            var result = fileDialog.ShowDialog();
            if (result.ToString() != string.Empty)
            {
                txtPath.Text = fileDialog.FileName;

                GetData(txtPath.Text);
            }

        }


        public void CollectData(string SearchItemOne, string SearchItemTwo)
        {

            //Everything to do with Selenium Searching.
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.bbc.co.uk/food");

            IWebElement element = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[1]/div/div/form/div/input[1]"));
            driver.Manage().Window.Maximize();
            // input food items
            element.SendKeys(SearchItemOne + " And " + SearchItemTwo);
            element.Submit();

            //Get the search results panel that contains the link for each result.
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[3]/div/div[4]"));
            //Get all the links only contained within the search result panel.
            IWebElement topMealClick = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[3]/div/div[4]/div[3]/div/div[1]/a"));
            topMealClick.Click();
            //Submit top meal
            IWebElement topMealName = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[4]/div/div[1]/div/div[1]/div[1]/div/h1"));
            Topmeal = topMealName.Text;
            driver.Close();
        }

        //method
        public void GetData(string path)
        {
            try
            {
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }



                ISheet sheet = hssfwb.GetSheet("Recipes");
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                rows.MoveNext();
                IRow HeaderRow = sheet.GetRow(0);

                //get each header
                foreach (ICell headerCell in HeaderRow)
                {
                    DataGridView.Columns.Add(new DataGridTextColumn { Header = headerCell.ToString(), Binding = new Binding(headerCell.ToString()) });
                }


                int rowCount = sheet.LastRowNum;
                for (int i = 1; i <= rowCount; i++)
                {
                    IRow curRow = sheet.GetRow(i);

                    var cellValue0 = curRow.GetCell(0).StringCellValue.Trim();
                    var cellValue1 = curRow.GetCell(1).StringCellValue.Trim();
                    var cellValue2 = curRow.GetCell(2).StringCellValue.Trim();
                    SearchItemOne = cellValue1;
                    SearchItemTwo = cellValue2;
                    CollectData(SearchItemOne, SearchItemTwo);
                    DataGridView.Items.Add(new { Name = cellValue0, IngredientOne = cellValue1, IngredientTwo = cellValue2, TopMeal = Topmeal });


                }
            }
            catch (Exception ex)
            {

            }


        }
    }
}