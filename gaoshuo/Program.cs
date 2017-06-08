using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace gaoshuo
{
    class Program
    {
        static void Main(string[] args)
        {
            ChromeDriver driver = new ChromeDriver();
            int pagesNum = 200;

            //login in
            driver.Navigate().GoToUrl("http://www.kanzhun.com/login/?ka=head-signin");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(600));
            //wait fields appear
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("username")));
            //enter credentials and login
            var usernameField = driver.FindElement(By.Id("username"));
            usernameField.SendKeys("18500285527");
            var pwdField = driver.FindElement(By.Id("password"));
            pwdField.SendKeys("gaoshuo123?");
            var loginBtn = driver.FindElement(By.XPath("//input[@type='submit']"));
            loginBtn.Click();
            //wait for a while
            Thread.Sleep(5000);

            List<RoleDetail> rdList = new List<RoleDetail>();
            try
            {
                for (int i = 12500; i <i+ pagesNum; i++)
                {

                    driver.Navigate().GoToUrl("http://www.kanzhun.com/gzxs" + i + ".html?ka=comsalary-detail-1");
                    //wait company name visible
                    wait.Until(ExpectedConditions.ElementExists(
                        By.XPath("/html/body/div[1]/section[2]/section/div/div[1]/a/div[1]/div[1]/span")));
                    //find category
                    string roleName;
                    try
                    {
                        roleName = getTextOfASpan(driver, "/html/body/div[1]/section[1]/section[1]/div[1]/span");
                    }
                    catch (Exception e)
                    {
                        continue;
                    }

                    //find company name
                    string companyName = getTextOfASpan(driver,
                        "/html/body/div[1]/section[2]/section/div/div[1]/a/div[1]/div[1]/span");
                    //find category
                    string category = getTextOfASpan(driver,
                        "/html/body/div[1]/section[2]/section/div/div[1]/a/div[1]/div[2]/span[1]");
                    //salary
                    string salary = getTextOfASpan(driver, "/html/body/div[1]/section[1]/section[1]/div[2]/div[1]/p[1]");
                    //base salary
                    string baseSalary = getTextOfASpan(driver, "/html/body/div[1]/section[1]/div[1]/div[3]/div[1]/span[2]");
                    //bonus
                    string bonus = getTextOfASpan(driver, "/html/body/div[1]/section[1]/div[1]/div[3]/div[2]/span[2]");
                    //support
                    string support = getTextOfASpan(driver, "/html/body/div[1]/section[1]/div[1]/div[3]/div[3]/span[2]");
                    //others
                    string saleCredit = getTextOfASpan(driver, "/html/body/div[1]/section[1]/div[1]/div[3]/div[4]/span[2]");
                    //others
                    string others = getTextOfASpan(driver, "/html/body/div[1]/section[1]/div[1]/div[3]/div[5]/span[2]");
                    RoleDetail rd = new RoleDetail()
                    {
                        BaseSalary = baseSalary,
                        Bonus = bonus,
                        Category = category,
                        Support = support,
                        SaleCredit = saleCredit,
                        Others = others,
                        CompanyName = companyName,
                        RoleName = roleName,
                        Salary = salary
                    };
                    rdList.Add(rd);
                }
            }
            finally
            {
                if (rdList.Count > 0)
                    SaveData(rdList);
            }

            //store to an excel
        }

        private static string getTextOfASpan(ChromeDriver driver, string xpath)
        {
            var Span =
                driver.FindElementByXPath("xpath");
            return Span.Text;
        }

        private static void SaveData(List<RoleDetail> rdList
            )
        {
            using (FileStream fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
                HSSFWorkbook wb = new HSSFWorkbook();
                ISheet sheet = wb.CreateSheet("sheet1");
                for (int i = 0; i < rdList.Count; i++)
                {
                    RoleDetail rd = rdList[i];
                    CreateRow(sheet, 0, rd.Category);
                    CreateRow(sheet, 1, rd.CompanyName);
                    CreateRow(sheet, 2, rd.RoleName);
                    CreateRow(sheet, 3, rd.Salary);
                    CreateRow(sheet, 4, rd.BaseSalary);
                    CreateRow(sheet, 5, rd.Bonus);
                    CreateRow(sheet, 6, rd.Support);
                    CreateRow(sheet, 7, rd.SaleCredit);
                    CreateRow(sheet, 8, rd.Others);
                }
                wb.Write(fs);
            }
        }

        private static void CreateRow(ISheet sheet, int index, string value)
        {
            IRow row = sheet.CreateRow(index);
            row.CreateCell(index).SetCellValue(value);
        }
    }

    class RoleDetail
    {
        public string CompanyName { get; set; }
        public string RoleName { get; set; }
        public string Category { get; set; }
        public string Salary { get; set; }
        public string BaseSalary { get; set; }
        public string Bonus { get; set; }
        public string Support { get; set; }
        public string SaleCredit { get; set; }
        public string Others { get; set; }

        public RoleDetail(string companyName, string roleName, string category, string salary, string baseSalary, string bonus, string support, string saleCredit, string others)
        {
            CompanyName = companyName;
            RoleName = roleName;
            Category = category;
            Salary = salary;
            BaseSalary = baseSalary;
            Bonus = bonus;
            Support = support;
            SaleCredit = saleCredit;
            Others = others;
        }

        public RoleDetail()
        {
        }
    }
}
