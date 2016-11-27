using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

namespace SDReport
{
	class Tools
	{
		public bool creatingError { get; }

		public int reqAmount { get; protected set; }
		public int techAmount { get; }
		public string[] technicians { get; }

		public Tools()
		{
			try
			{
				using (XmlReader reader = XmlReader.Create(@"ini.xml"))
				{
					reader.ReadToFollowing("ini");
					reader.ReadToFollowing("lastrequest");
					reqAmount = reader.ReadElementContentAsInt();

					reader.ReadToFollowing("technicians");
					reader.ReadToFollowing("amount");
					techAmount = reader.ReadElementContentAsInt();
					technicians = new string[techAmount];
					for (int i = 0; i < techAmount; i++)
					{
						reader.ReadToFollowing("technician");
						technicians[i] = reader.ReadElementContentAsString();
					}
				}
					creatingError = false;
					updateReqAmount();				
			}
			catch (Exception)
			{
				Console.WriteLine("Could not find \"ini.xml\" file or it is not valid.");
				creatingError = true;
			}
		}

		public void updateReqAmount()
		{
			int newReqAmount = reqAmount;
			int err = 0;
			int i = 0;
			Request req = new Request(reqAmount);
			do
			{
				i++;
				Console.Write($"Request #{reqAmount + i} ");
				req = new Request(reqAmount + i);
				if (req.sdp_status == "Failed")
				{
					Console.Write("Failed\n");
					err++;
				}
				else if (req.sdp_status == "Success")
				{
					Console.Write("Success\n");
					newReqAmount = req.workorderid;
					err = 0;
				}
				else
					err = 100;
			}
			while (err < 30);
			Console.Clear();

			try
			{
				XmlDocument doc = new XmlDocument();
				doc.Load(@"ini.xml");
				XmlElement ini = doc.DocumentElement;
				XmlNode lastreq = ini.FirstChild;
				lastreq.RemoveChild(lastreq.FirstChild);
				lastreq.AppendChild(doc.CreateTextNode(newReqAmount.ToString()));

				doc.Save(@"ini.xml");
				Console.WriteLine("\"ini.xml\" file has been changed");
			}
			catch (Exception)
			{
				Console.WriteLine("Could not find \"ini.xml\" file or it is not valid.");
				Console.WriteLine("ini.xml file has not been changed");
			}
			reqAmount = newReqAmount;
			Console.WriteLine($"Last request is #{reqAmount}");
		}

		public void weeklyResolvedRequests()
		{
			long week = 86400 * 7 * 1000;

			int[] made = new int[techAmount];
			for (int j = 0; j < techAmount; j++)
				made[j] = 0;

			Request req = new Request(reqAmount);
			long lTime = req.createdtime;

			int i = reqAmount;
			do
			{
				req = new Request(i);
				if (req.sdp_status == "Failed")
				{
					Console.WriteLine($"Checked request #{i} - Does not exist");
					i--;
				}
				else if (req.status != "Выполнено")
				{
					Console.WriteLine($"Checked request #{i} - Pending");
					i--;
				}
				else if (req.resolvedtime < lTime - week - 32400000) // 32400000 = 9 hours
				{
					Console.WriteLine($"Checked request #{i} - Resolved");
					i--;
				}
				else if (req.sdp_status == "Success" && req.resolvedtime > lTime - week - 32400000)
				{
					for (int j = 0; j < techAmount; j++)
					{
						if (req.technician == technicians[j])
							made[j]++;
					}
					Console.WriteLine($"Checked request #{i} - Recently resolved");
					i--;
				}
			}
			while (req.createdtime > lTime - week * 5 || req.sdp_status == "Failed");   // checking for last 5 weeks

			Console.Clear();
			for (int j = 0; j < techAmount; j++)
			{
				Console.WriteLine($"{technicians[j]}\t{made[j]} requests");
			}

		}

		public void createTsvReport()
		{
			long week = 86400 * 7 * 1000;

			int[] made = new int[techAmount];
			int[] pending = new int[techAmount];
			for (int j = 0; j < techAmount; j++)
			{
				made[j] = 0;
				pending[j] = 0;
			}

			Request req = new Request(reqAmount);
			long lTime = req.createdtime;

			int i = reqAmount;

			List<Request>[] resolvedList = new List<Request>[techAmount];
			List<Request>[] pendingList = new List<Request>[techAmount];

			for (int l = 0; l < techAmount; l++)
			{
				resolvedList[l] = new List<Request>();
				pendingList[l] = new List<Request>();
			}

			do
			{
				req = new Request(i);
				if (req.sdp_status == "Failed")
				{
					Console.WriteLine($"Checked request #{i} - Does not exist");
					i--;
				}
				else if (req.status != "Выполнено")
				{
					if (req.status == "Зарегистрирована" || req.status == "В ожидании")
					{
						for (int j = 0; j < techAmount; j++)
						{
							if (req.technician == technicians[j])
							{
								pending[j]++;
								pendingList[j].Add(req);
							}
						}
					}
					Console.WriteLine($"Checked request #{i} - Pending");
					i--;
				}
				else if (req.resolvedtime < lTime - week - 32400000) // 32400000 = 9 hours
				{
					Console.WriteLine($"Checked request #{i} - Resolved");
					i--;
				}
				else if (req.sdp_status == "Success" && req.resolvedtime > lTime - week - 32400000)
				{
					for (int j = 0; j < techAmount; j++)
					{
						if (req.technician == technicians[j])
						{
							made[j]++;
							resolvedList[j].Add(req);
						}
					}
					Console.WriteLine($"Checked request #{i} - Recently resolved");
					i--;
				}
			}
			while (req.createdtime > lTime - week * 5 || req.sdp_status == "Failed");   // checking for last 5 weeks

			string namep = DateTime.Now.ToString(@"dd/MM/yyyy hh-mm") + "p.tsv";
			string name = DateTime.Now.ToString(@"dd/MM/yyyy hh-mm") + ".tsv";
			string path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\AppData\Local\Temp\SDP\Reports\";

			Directory.CreateDirectory(path);

			// Resolved requests
			using (StreamWriter sw = new StreamWriter(File.Open(path + namep, FileMode.Create), Encoding.UTF32))
			{
				sw.WriteLine("\tID\tТема\tДата создания\tДата выполнения\tПотрачено времени\tПлощадка");
				for (int t = 0; t < techAmount; t++)
				{
					sw.WriteLine(technicians[t] + " (Выполнено " + made[t] + ")\t\t\t\t\t");
					foreach (Request rq in resolvedList[t])
					{
						sw.Write("\t");
						sw.Write(rq.workorderid + "\t");
						sw.Write(rq.subject + "\t");
						sw.Write(Request.longToDateTime(rq.createdtime) + "\t");
						sw.Write(Request.longToDateTime(rq.resolvedtime) + "\t");
						sw.Write(rq.timespentonreq + "\t");
						sw.Write(rq.area + "\n");
					}
				}
				sw.Close();
			}

			// Pending requests
			using (StreamWriter sw = new StreamWriter(File.Open(path + name, FileMode.Create), Encoding.UTF32))
			{
				sw.WriteLine("\tID\tТема\tАвтор заявки\tДата создания\tПлощадка\tПриоритет\tОписание");
				for (int t = 0; t < techAmount; t++)
				{
					sw.WriteLine(technicians[t] + " (Открыто " + pending[t] + ")\t\t\t\t\t\t");
					foreach (Request rq in pendingList[t])
					{
						sw.Write("\t");
						sw.Write(rq.workorderid + "\t");
						sw.Write(rq.subject + "\t");
						sw.Write(rq.requester + "\t");
						sw.Write(Request.longToDateTime(rq.createdtime) + "\t");
						sw.Write(rq.area + "\t");
						sw.Write(rq.priority + "\t");
						sw.Write(rq.getDescForExcel(90) + "\n");
					}
				}
				sw.Close();
			}
			
			Console.Clear();
			Console.WriteLine("Generated report file: " + path + namep);

			Application excel = new Application();

			Workbook wb = excel.Workbooks.Open(path + name);
			Worksheet ws1 = (Worksheet)wb.Worksheets[1];
			ws1.Name = "Открытые заявки";
			ws1.Columns.AutoFit();
			Range rng = ws1.Range["A1", "A2"].EntireColumn;
			rng.Font.Bold = true;
			rng = ws1.Range["A1", "B1"].EntireRow;
			rng.Font.Bold = true;

			Workbook wb2 = excel.Workbooks.Open(path + namep);
			Worksheet ws2 = (Worksheet)wb2.Worksheets[1];
			ws2.Name = "Выполненные за неделю";
			rng = ws2.Range["A1", "A2"].EntireColumn;
			rng.Font.Bold = true;
			rng = ws2.Range["A1", "B1"].EntireRow;
			rng.Font.Bold = true;
			ws2.Columns.AutoFit();

			ws2.Copy(Before:ws1);

			wb2.Close(SaveChanges:false);
			excel.Visible = true;
		}

	}
}
