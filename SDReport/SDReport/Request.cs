using System;
using System.Text;
using System.Net;
using System.Xml;
using System.Configuration;
using System.IO;

namespace SDReport
{
	class Request
	{
		public string sdp_status { get; set; }  // "Failed" if no such request; otherwise "Success"

		public int workorderid { get; set; }
		public string requester { get; set; }
		public string createdby { get; set; }
		public long createdtime { get; set; }
		public long resolvedtime { get; set; }
		public string shortdescription { get; set; }
		public string timespentonreq { get; set; }
		public string subject { get; set; }
		public string category { get; set; }
		public string subcategory { get; set; }
		public string technician { get; set; }
		public string status { get; set; }
		public string priority { get; set; }
		public string group { get; set; }
		public string description { get; set; }
		public string area { get; set; }

		public Request(int reqID)
		{
			WebClient wc = new WebClient();
			wc.Encoding = Encoding.UTF8;
			wc.Headers["Content-Type"] = "application/xml; charset=UTF-8";

			string reqStr = ConfigurationManager.AppSettings["SDP_PATH"] + "/request/" + reqID.ToString() + "?OPERATION_NAME=GET_REQUEST&TECHNICIAN_KEY=" + ConfigurationManager.AppSettings["SDP_API_KEY"];

			string xmlReqStr = null;
			try
			{
				xmlReqStr = wc.DownloadString(reqStr);
			}
			catch
			{
				xmlReqStr = "";
			}

			if (xmlReqStr.Length > 0)
			{

				try
				{
					using (XmlReader reader = XmlReader.Create(new StringReader(xmlReqStr)))
					{
						reader.ReadToFollowing("status");
						sdp_status = reader.ReadElementContentAsString();

						if (sdp_status == "Success")
						{
							reader.ReadToFollowing("value");
							workorderid = reader.ReadElementContentAsInt();
							reader.ReadToFollowing("value");
							requester = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							createdby = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							createdtime = reader.ReadElementContentAsLong();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							resolvedtime = reader.ReadElementContentAsLong();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							shortdescription = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							timespentonreq = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							subject = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							category = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							subcategory = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							technician = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							status = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							priority = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							reader.ReadToFollowing("value");
							group = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							description = reader.ReadElementContentAsString();
							reader.ReadToFollowing("value");
							area = reader.ReadElementContentAsString();
						}
						else
						{
							workorderid = reqID;
						}
					}
				}
				catch (Exception e)
				{
					Console.WriteLine(e.Message);
				}
			}
		}

		/// <summary>
		/// Formatting description string for excel output
		/// </summary>
		/// <param name="len">Length of line</param>
		/// <returns>String with special symbols</returns>
		public string getDescForExcel(int len)
		{
			string result = "";
			bool sw = false;
			for (int i = 0; i < shortdescription.Length; i++)
			{
				char cur = shortdescription[i];
				if (i != 0)
					if (i % len == 0)
					{
						sw = true;
					}
				if (sw && (cur == ' ' || cur == ',' || cur == ';'))
				{
					result += "\x000A\t\t\t\t\t\t\t";
					sw = false;
				}
				result += cur;
			}
			return result;
		}

		public static DateTime longToDateTime(long dateNumber)
		{
			long beginTicks = new DateTime(1970, 1, 1, 3, 0, 0, DateTimeKind.Utc).Ticks;
			return new DateTime(beginTicks + dateNumber * 10000);
		}

		public void consoleOutput()
		{
			if (sdp_status == "Success")
			{
				Console.WriteLine($"Request ID: {workorderid}");
				Console.WriteLine($"Requester: {requester}");
				Console.WriteLine($"Created by: {createdby}");
				Console.WriteLine($"Subject: {subject}");
				Console.WriteLine($"Category: {category}");
				Console.WriteLine($"Subcategory: {subcategory}");
				Console.WriteLine($"Short description: {shortdescription}");
				Console.WriteLine($"Technician: {technician}");
				Console.WriteLine($"Created time: {longToDateTime(createdtime)}");
				if (resolvedtime > 0)
				{
					Console.WriteLine($"Resolved time: {longToDateTime(resolvedtime)}");
					Console.WriteLine($"Time spent: {timespentonreq}");
				}
				Console.WriteLine($"Status: {status}");
				Console.WriteLine($"Priority: {priority}");
				Console.WriteLine($"Group: {group}");
				Console.WriteLine($"Area: {area}");
			}
			else
			{
				Console.WriteLine($"Request #{workorderid} does not exist.");
			}
		}

		public void consoleOutputShort()
		{
			if (sdp_status == "Success")
			{
				Console.WriteLine($"Request ID: {workorderid}");
				Console.WriteLine($"Subject: {subject}");
				Console.WriteLine($"Short description: {shortdescription}");
				Console.WriteLine($"Technician: {technician}");
				Console.WriteLine($"Created time: {longToDateTime(createdtime)}");
				if (resolvedtime > 0)
				{
					Console.WriteLine($"Resolved time: {longToDateTime(resolvedtime)}");
					Console.WriteLine($"Time spent: {timespentonreq}");
				}
				Console.WriteLine($"Status: {status}");
				Console.WriteLine($"Area: {area}");
			}
			else
			{
				Console.WriteLine($"Request #{workorderid} does not exist.");
			}
		}

	}
}
