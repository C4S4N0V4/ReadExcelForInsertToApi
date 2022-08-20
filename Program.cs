using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

class Program
{
    static void Main()
    {
        ReadExcelFile();
    }
    static void ReadExcelFile()
    {
        try
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(@"C:\\", false))
            {
                //create the object for workbook part  
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                StringBuilder excelResult = new StringBuilder();

                //using for each loop to get the sheet from the sheetcollection  
                foreach (Sheet thesheet in thesheetcollection)
                {
                    excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                    excelResult.AppendLine("----------------------------------------------- ");
                    //statement to get the worksheet object by using the sheet id  
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                    SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                    var count = 0;
                    IList<Client> list = new List<Client>();
                    foreach (Row thecurrentrow in thesheetdata)
                    {
                        if (count == 0)
                        {

                        }
                        else if (count > 0)
                        {
                            Client c = new Client();
                            foreach (Cell thecurrentcell in thecurrentrow)
                            {
                                //statement to take the integer value

                                string currentcellvalue = string.Empty;
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                if (thecurrentcell.CellReference.ToString().Contains("C"))
                                                {
                                                    c.Email = item.Text.InnerText ?? "";
                                                }
                                                else if (thecurrentcell.CellReference.ToString().Contains("S"))
                                                {
                                                    c.Name = item.Text.InnerText ?? "";
                                                }
                                                //code to take the string value  
                                                excelResult.Append(item.Text.Text + " ");
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentcellvalue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentcellvalue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                            }
                            list.Add(c);
                        }
                        count = count + 1;
                        excelResult.AppendLine();
                    }
                    foreach (var client in list)
                    {
                        Subscrib(client);
                    }
                    excelResult.Append("");
                    Console.WriteLine(excelResult.ToString());
                    Console.ReadLine();
                }
            }
        }
        catch (Exception)
        {

        }
    }
    static void Subscrib(Client c)
    {
        var client = new RestClient("https://localhost/subscribe?ResponseFormat=JSON&ListID=17670&EmailAddress=" + c.Email + "&CustomField1=" + c.Name );
        client.Timeout = -1;
        //client.Proxy = new WebProxy("https://177.103.222.139");
        var request = new RestRequest(Method.POST);
        request.AddHeader("Cookie", "PHPSESSID=acp5b3qkb09inni9n96135imm5");
        request.AddParameter("text/plain", null, ParameterType.RequestBody);
        IRestResponse response = client.Execute(request);
        if (response.StatusCode.Equals(HttpStatusCode.OK) || response.StatusCode.Equals(HttpStatusCode.Created))
        {
            Console.WriteLine("Created");
            Console.WriteLine(response.Content);
        }
        else
        {

            Console.WriteLine(response.Content);
        }
    }
    public class Client
    {
        public string Email { get; set; }
        public string Name { get; set; }
    }
}
