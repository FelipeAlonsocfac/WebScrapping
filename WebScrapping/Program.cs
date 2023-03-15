using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebScrapping;

string companiesXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[1]/span/a";
string layoutsXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[2]/span";
string datesXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[3]/span";
string industriesXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[4]/span";
string hQSXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[5]/span";
string sourcesXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[6]/span";
string companiesStatusXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[7]/span";
string notesXpath = "//*[@id=\"ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31\"]/div/div/div/div[2]/div/table/tbody/tr/td[8]/span";

//*[@id="ContentItemDiv-74bf8cbe-d89d-4450-b3fb-7858bf021c31"]/div/div/div/div[2]/div/table/tbody/tr[503]/td[3]/span
IWebDriver driver = new ChromeDriver();
driver.Navigate().GoToUrl("https://infogram.com/crunchbase-layoffs-tracker-1h8n6m3ogl3xz4x");
var js = (IJavaScriptExecutor)driver;
WaitForElementToExist(driver, companiesXpath);
var CompaniesNbr = 0;
var companyElement = driver.FindElements(By.XPath(companiesXpath));
var newCompaniesNbr = companyElement.Count;
while (CompaniesNbr != newCompaniesNbr)
{
    var lastCompany = companyElement[companyElement.Count - 1];
    CompaniesNbr = newCompaniesNbr;
    js.ExecuteScript("arguments[0].scrollIntoView();", lastCompany);
    Thread.Sleep(1);
    companyElement = driver.FindElements(By.XPath(companiesXpath));
    newCompaniesNbr = companyElement.Count;
}
var companies = companyElement.Select(x => x.Text).ToList();
var layouts = driver.FindElements(By.XPath(layoutsXpath)).Select(x => x.Text).ToList();
var dates = driver.FindElements(By.XPath(datesXpath)).Select(x => x.Text).ToList();
var industries = driver.FindElements(By.XPath(industriesXpath)).Select(x => x.Text).ToList();
var hQS = driver.FindElements(By.XPath(hQSXpath)).Select(x => x.Text).ToList();
var sources = driver.FindElements(By.XPath(sourcesXpath)).Select(x => x.Text).ToList();
var companiesStatus = driver.FindElements(By.XPath(companiesStatusXpath)).Select(x => x.Text).ToList();
var notes = driver.FindElements(By.XPath(notesXpath)).Select(x => x.Text).ToList();
List<Layoff> layoffs = new();
for (int i = 0; i < companies.Count() - 1; i++)
{
    layoffs.Add(new Layoff
    {
        Company = companies[i],
        NumberOfLayoffs = layouts[i],
        ReportedDate = dates[i],
        Industry = industries[i],
        HQ = hQS[i],
        Source = sources[i],
        CompanyStatus = companiesStatus[i],
        Notes = notes[i]
    });
    //Console.WriteLine($"{i + 1}   ///   {companies[i].Text}  ///  {layouts[i].Text}  ///  {dates[i].Text}   ///   {industries[i].Text}   ///   {hQS[i].Text}   ///   {sources[i].Text}   ///   {companiesStatus[i].Text}   ///   {notes[i].Text}");
}

WriteToExcel(layoffs);


static void WaitForElementToExist(IWebDriver driver, string xpath)
{
    while (driver.FindElements(By.XPath(xpath)).Count < 1)
    {
        Thread.Sleep(1);
    }
}


static void WriteToExcel(List<Layoff> layoffs) {

    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    var file = new FileInfo(@"C:\Users\xfeli\OneDrive\Escritorio\ExcelPrueba.xlsx");

    using var package = new ExcelPackage(file);
    var ws = package.Workbook.Worksheets.Add("CrunchBaseReport");

    var range = ws.Cells["A1"].LoadFromCollection(layoffs, true);
    range.AutoFitColumns();
    package.Save();
    package.Dispose();
}