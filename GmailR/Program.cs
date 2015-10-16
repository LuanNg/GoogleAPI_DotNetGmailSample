using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

//Now 2015/10/16
//Licence is same "Gmail API .NET Quickstart"

//Nuget
//PM> Install-Package Google.Apis.Gmail.v1

// ref add Microsoft.VisualBasic;

namespace GmailR
{
	internal class Program
	{
		//Gmail Access Authority
		private static string[] Scopes = { GmailService.Scope.GmailReadonly };

		private static string ApplicationName = "Gmail API .NET Quickstart";

		//Register Project and set ClientID and ClinetSecret in https://console.developers.google.com/project
		private const string CID = "YOUR CLINET ID";

		private const string CSECRET = "YOUR CLIENT SECRET";

		private static void Main(string[] args)
		{
			UserCredential credential;

			//Save Token Response Path
			string credPath = Environment.CurrentDirectory + @"\";
			credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart");

			credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
				new ClientSecrets { ClientId = CID, ClientSecret = CSECRET },
				Scopes,
				"user", // I don't know 
				CancellationToken.None,
				new FileDataStore(credPath, true)).Result;

			Console.WriteLine("Credential file saved to: " + credPath);

			// Create Gmail API service.
			var service = new GmailService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});

			// Define parameters of request.
			var requestLabel = service.Users.Labels.List("me");

			var requestLabelList = new List<string>();
			var requestLabelNameList = new List<string>() { "unread", "CATEGORY_UPDATES" };

			// List labels.
			var labels = requestLabel.Execute().Labels;

			Console.WriteLine("Labels:");
			if (labels != null && labels.Count > 0)
			{
				foreach (var labelItem in labels)
				{
					Console.WriteLine("{0}", labelItem.Name);

					//Getting wanted Label
					if (requestLabelNameList.Contains(labelItem.Name)) requestLabelList.Add(labelItem.Id);
				}
			}

			//GetMessage
			//Me means myself
			var request = service.Users.Messages.List("me");

			request.LabelIds = requestLabelList;

			//Get maximum message
			request.MaxResults = 4;

			//plain text query
			//request.Q = "is:unread category:updates is:important";

			ListMessagesResponse respM = request.Execute();
			if (respM.Messages != null)
			{
				foreach (var m in respM.Messages)
				{
					var id = m.Id;
					var getReq = new UsersResource.MessagesResource.GetRequest(service, "me", id);

					//Get Message from id
					var data = getReq.Execute();

					//Output Message's Label
					foreach (var l in data.LabelIds)
						Console.Write(l + " ");
					Console.WriteLine();

					var heads = data.Payload.Headers;

					/*
					//output full header info
					var hCount = 1;
					foreach(var head in heads)
					{
						Console.WriteLine("Head {0} {1} {2}", hCount, head.Name, head.Value);
						hCount++;
					}
					 */

					var subject = heads.FirstOrDefault(x => x.Name == "Subject")?.Value ?? "NotSubject";
					Console.WriteLine(subject);

					var body = GetMimeString(data.Payload);
					var fix = Strings.Replace(body, subject, "", 1);
					Console.WriteLine(fix);
					Console.WriteLine("++++++++++++++++++++++++++++++++++++");
				}
			}

			Console.Read();
		}

		public static string GetMimeString(MessagePart Parts)
		{
			var Body = "";

			if (Parts.Parts != null)
			{
				foreach (var part in Parts.Parts)
				{
					Body = string.Format("{0}\n{1}", Body, GetMimeString(part));
				}
			}
			else if (Parts.Body.Data != null && Parts.Body.AttachmentId == null) // && Parts.MimeType == "text/plain")
			{
				string codedBody = Parts.Body.Data.Replace("-", "+");
				codedBody = codedBody.Replace("_", "/");
				byte[] data = Convert.FromBase64String(codedBody);
				Body = Encoding.UTF8.GetString(data);
			}
			return Body;
		}
	}
}