using System;
using System.Text;
using System.Net;
using System.Xml;
using System.Configuration;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.NetworkInformation;

namespace SDReport
{
	class Program
	{
		static void Main(string[] args)
		{
			Ping ping = new Ping();
			IPAddress ip = new IPAddress(134744072);
			PingOptions options = new PingOptions();
			options.DontFragment = true;
			options.Ttl = 57;
			string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
			byte[] buffer = Encoding.ASCII.GetBytes(data);

			try {
				PingReply reply = ping.Send(ip, 3000, buffer, options);
				if (reply.Status == IPStatus.Success)
				{
					Console.WriteLine("Ping help.citysystems.su - Success");

					Tools t = new Tools();
					if (!t.creatingError)
					{
						t.createTsvReport();
					}
					else
					{
						Console.ReadKey();
					}
				}
				else
				{
					Console.WriteLine("Ping help.citysystems.su - Failed");
					Console.ReadKey();
				}
			}
			catch (Exception)
			{
				Console.WriteLine("Ping help.citysystems.su - Failed");
				Console.ReadKey();
			}
		}
	}	

}
