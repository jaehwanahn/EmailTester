using ARSoft.Tools.Net.Dns;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace EmailTester
{
    /// <summary>
    /// To check if an email address is valid.
    /// This is programmatic way to perform (nslookup, telnet x.x.x.x 25, etc)
    /// If a mail server has catch-all policy, it returns always 250. (meaning it always going to tell an email address is valid)
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            p.GenerateReport();
        }

        private string GetDomain(string emailAddress)
        {
            MailAddress address = new MailAddress(emailAddress);
            return address.Host;
        }

        // The email addresses in CSV file are seperated by next line.
        // Place the CSV file where exe file is located. ex.) bin/debug/
        private List<string> ReadEmailAddresses()
        {
            List<string> emailAddresses = new List<string>();
            using (var reader = new StreamReader("emails.csv"))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    emailAddresses.Add(line);
                }
            }
            return emailAddresses;
        }

        //https://stackoverflow.com/questions/2669841/how-to-get-mx-records-for-a-dns-name-with-system-net-dns
        private string GetDNS(string targetDomain)
        {
            var resolver = new DnsStubResolver();
            var records = resolver.Resolve<MxRecord>(targetDomain, RecordType.Mx);

            List<string> mxRecords = new List<string>();

            if (!string.IsNullOrEmpty(records.ToString()))
            {
                foreach (var record in records)
                {
                    string dummy = record.ExchangeDomainName?.ToString();
                    mxRecords.Add(dummy.Substring(0, dummy.Length - 1));
                }
            }
            else
            {
                return "";
            }

            return mxRecords[0];
        }

        //https://www.c-sharpcorner.com/UploadFile/kirtan007/check-if-email-address-really-exist-or-not-using-C-Sharp/
        private string TestEmail (string mxRecord, string targetEmail)
        {
            string response = "";
            TcpClient tClient = new TcpClient(mxRecord, 25);
            string CRLF = "\r\n";
            byte[] dataBuffer;
            string ResponseString;
            NetworkStream netStream = tClient.GetStream();
            StreamReader reader = new StreamReader(netStream);
            ResponseString = reader.ReadLine();
            /* Perform HELO to SMTP Server and get Response */
            dataBuffer = BytesFromString("EHLO Hi" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            ResponseString = reader.ReadLine();
            dataBuffer = BytesFromString("MAIL FROM:<test@test.com>" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            ResponseString = reader.ReadLine();
            /* Read Response of the RCPT TO Message to know from mail server if it exist or not */
            string rcpt = targetEmail;
            dataBuffer = BytesFromString("RCPT TO:<" + rcpt + ">" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            ResponseString = reader.ReadLine();
            if (GetResponseCode(ResponseString) != 250)
            {
                response = "The Address Does not Exist";
                //Console.WriteLine(ResponseString);
                //Console.WriteLine("The Address Does not Exist !");
                //Console.WriteLine("Original Error from Smtp Server : " + ResponseString);
            }
            else
            {
                response = "The Address Exists";
                //Console.WriteLine(ResponseString);
                //Console.WriteLine("The Address Exists!");
            }
            /* QUITE CONNECTION */
            dataBuffer = BytesFromString("QUITE" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            tClient.Close();
            return response;
        }

        private byte[] BytesFromString(string str)
        {
            return Encoding.ASCII.GetBytes(str);
        }
        private int GetResponseCode(string ResponseString)
        {
            return int.Parse(ResponseString.Substring(0, 3));
        }

        private List<Result> GetResult()
        {
            List<string> emails = ReadEmailAddresses();
            Result r;
            List<Result> results = new List<Result>();
            string targetDomain = "";
            string mxRecord = "";
            string result = "";

            for (int i = 0; i < emails.Count; i++)
            {
                targetDomain = GetDomain(emails[i]);
                mxRecord = GetDNS(targetDomain);
                if (string.IsNullOrEmpty(mxRecord))
                    continue;
                result = TestEmail(mxRecord, emails[i]);
                r = new Result(emails[i], result);
                results.Add(r);
            }
            return results;
        }

        //http://csharp.net-informations.com/excel/csharp-create-excel.htm
        private void GenerateReport()
        {
            List<Result> results = GetResult();

            Application xlApp = new Application();

            Workbook xlWorkBook;
            Worksheet xlWorkSheet;

            object misValue = Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Email Address";
            xlWorkSheet.Cells[1, 2] = "Result";

            int j = 2;
            for (int i = 0; i < results.Count; i++)
            {                
                xlWorkSheet.Cells[j, 1] = results[i].emailAddress;
                xlWorkSheet.Cells[j, 2] = results[i].message;
                j++;
            }

            // Need to specify full path.
            xlWorkBook.SaveAs(@"D:\C#_Projects\EmailTester\testresult.xls", XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }
    }
}
