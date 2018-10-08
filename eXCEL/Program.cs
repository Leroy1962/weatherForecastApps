using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace eXCEL
{
	class Program
	{
		static void Main(string[] args)
		{
			Excel.Application xlApp = new Excel.Application();
			xlApp.Visible = true;
			Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\User\Desktop\10.xlsx");
			//xlApp.Visible = false;
			Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
			Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
			Excel.Range xlRange = xlWorksheet.UsedRange;
			Excel.Range xlRange2 = xlWorksheet2.UsedRange;
			int rowCount = xlRange.Rows.Count;

			int rowCount2 = rowCount;
			string ipAdres = "46.16.15.23";
			Console.WriteLine(ipAdres);
			string[] country = new string[rowCount + 1];
			List<string> countryIP = new List<string>();
			//xlRange2.Delete();
			//xlRange2.Columns[1].AutoFit();
			//xlRange2.Columns[2].AutoFit();

			HttpClient client = new HttpClient();
			
			var result = client.GetStringAsync($"http://freegeoip.net/json/{ipAdres.Trim()}");
			var json = result.GetAwaiter().GetResult();
			JObject o = JObject.Parse(json);

			string name = (string)o["country_name"];

			for (int i = 1; i <= rowCount; i++)
			{

				if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
				{
					
					ipAdres = xlRange.Cells[i, 1].Value2.ToString();
					result = client.GetStringAsync($"http://freegeoip.net/json/{ipAdres.Trim()}");
					json = result.GetAwaiter().GetResult();
					o = JObject.Parse(json);
					name = (string)o["country_name"];
					xlRange.Cells[i, 2].Value = name;
					//xlRange2.Cells[i, 1].Value = name;
					//xlRange2.Cells[i, 3].Value = name;
					countryIP.Add(name);
					Console.WriteLine("Country: " + name);
				}
				
				
				/*for (int i = 1; i <= rowCount; i++)
				{

					if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
					{
						GeoService.GeoIPService service1 = new GeoService.GeoIPService();
						GeoService.GeoIP output1;
						ipAdres = xlRange.Cells[i, 1].Value2.ToString();
						Console.WriteLine(ipAdres);
						output1 = service1.GetGeoIP(ipAdres.Trim());
						Console.WriteLine("Country: " + output1.CountryName);
						xlRange.Cells[i, 2].Value = output1.CountryName;
						xlRange2.Cells[i, 1].Value = output1.CountryName;
						xlRange2.Cells[i, 3].Value = output1.CountryName;
						country[i] = output1.CountryName;
						countryIP.Add(output1.CountryName);
					}*/

			}
			int pi = 1;
			foreach (var val in countryIP.Distinct())
			{
				Console.WriteLine(val + " - " + countryIP.Where(x => x == val).Count() + " раз");
				xlRange2.Cells[pi, 1].Value = val;
				xlRange2.Cells[pi, 2].Value = countryIP.Where(x => x == val).Count();
				pi++;

			}
			/*string sd = "=COUNTIF(A1:A20,C1)";
			int count1 = 1;
			for (int i = 1; i <= rowCount; i++)
			{
				xlRange2.Cells[i, 2].Formula = sd;
				xlRange.Cells[i, 3] = xlRange2.Cells[i, 2];
				xlRange2.Cells[i, 2] = xlRange2.Cells[i, 2];
				count1++;
				sd = "=COUNTIF(A1:A20,C" + count1 + ")";
			}*/

			//xlRange2.RemoveDuplicates(1);
			xlRange2.Sort(xlRange2.Columns[2], Excel.XlSortOrder.xlDescending);
			//((Range)xlRange2.Columns[3]).Clear();
			//int df = xlRange2.Rows.Count;

			//Console.WriteLine(df);

			GC.Collect();
			GC.WaitForPendingFinalizers();

			//release com objects to fully kill excel process from running in the background
			Marshal.ReleaseComObject(xlRange);
			Marshal.ReleaseComObject(xlRange2);
			Marshal.ReleaseComObject(xlWorksheet);

			//close and release
			xlWorkbook.Close();
			Marshal.ReleaseComObject(xlWorkbook);

			//quit and release
			xlApp.Quit();
			Marshal.ReleaseComObject(xlApp);

			Console.ReadLine();
		}

	}
}
