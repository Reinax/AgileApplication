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
                //Method for Excel create method to input excel data into datagrid.
                GetData(txtPath.Text);




                //Everything to do with Selenium Searching.
                IWebDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl("https://www.bbc.co.uk/food");

                IWebElement element = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[1]/div/div/form/div/input[1]"));
                driver.Manage().Window.Maximize();
                // input food items
                element.SendKeys("Cheese and Milk");
                element.Submit();

                //Get the search results panel that contains the link for each result.
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[3]/div/div[4]"));
                //Get all the links only contained within the search result panel.
                IWebElement topMealClick = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[3]/div/div[4]/div[3]/div/div[1]/a"));
                topMealClick.Click();
                //Submit top meal
                IWebElement topMealName = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[4]/div/div[1]/div/div[1]/div[1]/div/h1"));
                InputCheck.Text = topMealName.Text;



                // Print the text for every link in the search results.
                /*for (int i = 0; i < searchResults.Count; i++)
                {
                    DataGridView.Items.Add(searchResults[i].Text);
                    
                }
                */


            }

        }

        //method
        public void GetData(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            DataGridView.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new Binding("Name") });
            DataGridView.Columns.Add(new DataGridTextColumn { Header = "Ingredient 1", Binding = new Binding("Ingredient 1") });
            DataGridView.Columns.Add(new DataGridTextColumn { Header = "Ingredient 2", Binding = new Binding("Ingredient 2") });
            DataGridView.Columns.Add(new DataGridTextColumn { Header = "Top Meal", Binding = new Binding("Top Meal") });
            ISheet sheet = hssfwb.GetSheet("Recipes");
            for (int row = 0; row <= sheet.LastRowNum+1; row++)
            {
                
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(1).StringCellValue));
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(2).StringCellValue));
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(3).StringCellValue));
                    //DataGridView.Items.Add(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                    //DataGridView.Items.Add(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(1).StringCellValue));
                    //DataGridView.Items.Add(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(2).StringCellValue));
                    //DataGridView.Items.Add(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(3).StringCellValue));


                }
            }
        }

    }
}