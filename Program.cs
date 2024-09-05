using HtmlAgilityPack;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SimpleWebScraper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

class Program
{
    static void Main()
    {
        ChromeOptions options = new ChromeOptions();
        // options.AddArgument("--headless"); 
        options.AddArgument("--disable-gpu");

        using (IWebDriver driver = new ChromeDriver())
        {
            try
            {
                driver.Navigate().GoToUrl("https://app.socio.events/MjkyNjQ/Attendees/401456");

              
                driver.FindElement(By.Id("email")).SendKeys("kevin@thriveagency.com");
                driver.FindElement(By.Id("continue-button")).Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.FindElement(By.Id("password-input")).SendKeys("Thrive2024");
                driver.FindElement(By.Id("login-button")).Click();

                
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(180));
                wait.Until(drv => drv.FindElements(By.XPath("//li[contains(@class, 'MuiListItem-root')]")).Count > 0);

                var attendees = new List<Attendees>();
                bool moreData = true;
                int previousCount = 0;

                while (moreData)
                {
                    
                    ExtractAttendees(driver, attendees);

                 
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollBy(0, document.body.scrollHeight);");
                    Console.WriteLine("Scrolling down...");
                    Thread.Sleep(5000); 

                   
                    ExtractAttendees(driver, attendees);

                  
                    int currentCount = attendees.Count;
                    if (currentCount == previousCount)
                    {
                        moreData = false;  
                        Console.WriteLine("No more data to load.");
                    }
                    else
                    {
                        previousCount = currentCount;
                    }
                }

               
                InsertDataIntoCsv(attendees);
                string json = JsonConvert.SerializeObject(attendees, Formatting.Indented);
                Console.WriteLine(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

 
    static void ExtractAttendees(IWebDriver driver, List<Attendees> attendees)
    {
        var pageSource = driver.PageSource;
        var document = new HtmlDocument();
        document.LoadHtml(pageSource);

        var attendeesHTMLElements = document.DocumentNode.SelectNodes("//li[contains(@class, 'MuiListItem-root')]");

        if (attendeesHTMLElements != null)
        {
            foreach (var element in attendeesHTMLElements)
            {
                var personNameNode = element.SelectSingleNode(".//span[@data-testid='attendee-list-item-full-name']");
                var titleNode = element.SelectSingleNode(".//span[@data-testid='attendee-list-item-title']");
                var companyNameNode = element.SelectSingleNode(".//span[@data-testid='attendee-list-item-company']");

                var personName = HtmlEntity.DeEntitize(personNameNode?.InnerText ?? "N/A");
                var title = HtmlEntity.DeEntitize(titleNode?.InnerText ?? "N/A");
                var companyName = HtmlEntity.DeEntitize(companyNameNode?.InnerText ?? "N/A");

                var attendee = new Attendees { PersonName = personName, Title = title, CompanyName = companyName };
                if (!attendees.Any(a => a.PersonName == attendee.PersonName && a.CompanyName == attendee.CompanyName))
                {
                    attendees.Add(attendee);
                }
            }
        }
    }

   
    static void InsertDataIntoCsv(List<Attendees> attendees)
    {
       
        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        string fileName = $"Attendees4-{currentDate}.csv";

        
        string desiredPath = @"D:\";
        string csvFilePath = Path.Combine(desiredPath, fileName);

        try
        {
            
            using (var writer = new StreamWriter(csvFilePath))
            {
                writer.WriteLine("Person Name,Title,Company Name");
                foreach (var attendee in attendees)
                {
                    writer.WriteLine($"\"{attendee.PersonName}\",\"{attendee.Title}\",\"{attendee.CompanyName}\"");
                }
            }

            
            Console.WriteLine($"Data successfully saved to: {csvFilePath}");

           
            System.Diagnostics.Process.Start("explorer.exe", desiredPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while saving the file: {ex.Message}");
        }
    }
}

namespace SimpleWebScraper
{
    class Attendees
    {
        public string PersonName { get; set; }
        public string Title { get; set; }
        public string CompanyName { get; set; }
    }
}