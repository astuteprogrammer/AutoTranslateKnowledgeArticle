using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace KnowledgeArticleTranslationSolution
{
	/// <summary>
	/// Helper class that handles the actual translation
	/// </summary>
	internal static class MicrosoftTranslateHelper
	{
		/// <summary>
		/// This method translate the content from source language to destination language.
		/// </summary>
		/// <param name="articleContent"> Input text to translate</param>
		/// <param name="sourceLanguageLCID"> Source language ID</param>
		/// <param name="targetLanguageLCID"> Destinatino language ID</param>
		/// <param name="apiKey"> Microsoft Translator API Key</param>
		/// <param name="clientId"> Micorsoft Translator Client Application ID</param>
		/// <returns> Translated text</returns>
		internal static string Translate(string articleContent, string sourceLanguageLCID, string targetLanguageLCID, string apiKey, string clientId)
		{
			AdmAccessToken admToken;
			string headerValue = "";
			AdmAuthentication admAuth = new AdmAuthentication(clientId, apiKey);
			admToken = admAuth.GetAccessToken();
			headerValue = "Bearer " + admToken.access_token;

			WebResponse response = null;
			string result = "";
			try
			{
				string uri = "http://api.microsofttranslator.com/v2/Http.svc/TranslateArray";
				string body = "<TranslateArrayRequest>" +
								 "<AppId />" +
								 "<From>{0}</From>" +
								 "<Options>" +
									" <Category xmlns=\"http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2\" />" +
									 "<ContentType xmlns=\"http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2\">{1}</ContentType>" +
									 "<ReservedFlags xmlns=\"http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2\" />" +
									 "<State xmlns=\"http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2\" />" +
									 "<Uri xmlns=\"http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2\" />" +
									 "<User xmlns=\"http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2\" />" +
								 "</Options>" +
								 "<Texts>" +
									"<string xmlns=\"http://schemas.microsoft.com/2003/10/Serialization/Arrays\">{2}</string>" +
								 "</Texts>" +
								 "<To>{3}</To>" +
							  "</TranslateArrayRequest>";
				string reqBody = string.Format(body, sourceLanguageLCID, "text/html", articleContent.Replace("<","&lt;").Replace(">","&gt;"), targetLanguageLCID);

				// create the request
				HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
				httpWebRequest.ContentType = "text/xml";
				httpWebRequest.Method = "POST";

				using (System.IO.Stream stream = httpWebRequest.GetRequestStream())
				{
					byte[] arrBytes = System.Text.Encoding.UTF8.GetBytes(reqBody);
					stream.Write(arrBytes, 0, arrBytes.Length);
				}

				httpWebRequest.Headers.Add("Authorization", headerValue);
				response = httpWebRequest.GetResponse();

				using (Stream stream = response.GetResponseStream())
				{
					using (StreamReader rdr = new StreamReader(stream, System.Text.Encoding.UTF8))
					{
						// Deserialize the response
						string strResponse = rdr.ReadToEnd();
						Console.WriteLine("Result of translate array method is:");
						XDocument doc = XDocument.Parse(@strResponse);
						XNamespace ns = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2";
						int soureceTextCounter = 0;
						foreach (XElement xe in doc.Descendants(ns + "TranslateArrayResponse"))
						{
							foreach (var node in xe.Elements(ns + "TranslatedText"))
							{
								result += node.Value;
							}
							soureceTextCounter++;
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
			finally
			{
				if (response != null)
				{
					response.Close();
					response = null;
				}
			}
			return result; ;
		}

		/// <summary>
		/// This method returns Translator API settings object 
		/// </summary>
		/// <param name="service">IOrganizationService object</param>
		/// <param name="tracingService">ITracingService object</param>
		/// <returns>TranslateApiSettings object</returns>
		internal static TranslateApiSettings GetTranslatorAPISettings(IOrganizationService service, ITracingService tracingService)
		{
			EntityCollection retrieved;
			TranslateApiSettings apiSettings = new TranslateApiSettings();
			var cols = new ColumnSet();
			cols.AddColumns(new string[] { "msdyn_apikey", "msdyn_name", "msdyn_isautotranslationenabled" });
			try
			{
				var query = new QueryExpression
				{
					ColumnSet = cols,
					EntityName = "msdyn_automatickmtranslationsetting",
				};
				retrieved = service.RetrieveMultiple(query);
				if (retrieved.Entities.Count == 1)
				{
					if (retrieved.Entities[0].Attributes.Contains("msdyn_apikey"))
					{
						apiSettings.ClientSecret = (string)retrieved.Entities[0].Attributes["msdyn_apikey"];
					}

					if (retrieved.Entities[0].Attributes.Contains("msdyn_name"))
					{
						apiSettings.ClientID = (string)retrieved.Entities[0].Attributes["msdyn_name"];
					}

					if (retrieved.Entities[0].Attributes.Contains("msdyn_isautotranslationenabled"))
					{
						apiSettings.Isautotranslationenabled = (bool)retrieved.Entities[0].Attributes["msdyn_isautotranslationenabled"];
					}
				}
			}
			catch (Exception ex)
			{
				tracingService.Trace("GetTranslatorAPISettings: {0}", ex.Message);
			}
			return apiSettings;
		}

		/// <summary>
		/// This method returns record of automatickmtranslationsetting custom entity
		/// </summary>
		/// <param name="service">IOrganizationService object</param>
		/// <param name="tracingService">ITracingService object</param>
		/// <returns>returns count of records</returns>
		internal static int GetTranslatorAPISettingsRecordsCount(IOrganizationService service, ITracingService tracingService)
		{
			EntityCollection retrieved;
			int recordsCount = 0;
			var cols = new ColumnSet();
			cols.AddColumns(new string[] { "msdyn_apikey", "msdyn_name", "msdyn_isautotranslationenabled" });
			try
			{
				var query = new QueryExpression
				{
					ColumnSet = cols,
					EntityName = "msdyn_automatickmtranslationsetting",
				};
				retrieved = service.RetrieveMultiple(query);
				recordsCount = retrieved.Entities.Count;
			}
			catch (Exception ex)
			{
				tracingService.Trace("GetTranslatorAPISettingsRecordsCount: {0}", ex.Message);
			}
			return recordsCount;
		}

		/// <summary>
		/// This method returns the web resoruce file name based on Org and User language.
		/// </summary>
		/// <param name="organizationService"> IOrganizationService object</param>
		/// <param name="context">IPluginExecutionContext object</param>
		/// <returns>returns web resource file name</returns>
		internal static string GetLanguageResourceFile(IOrganizationService organizationService, IPluginExecutionContext context)
		{
			int OrgLanguage = RetrieveOrganizationBaseLanguageCode(organizationService);
			int UserLanguage = RetrieveUserUILanguageCode(organizationService, context.InitiatingUserId);
			String fallBackResourceFile = "";
			switch (OrgLanguage)
			{
				case 1033:
					fallBackResourceFile = "msdyn_localizedString.en_US";
					break;
				case 1041:
					fallBackResourceFile = "msdyn_localizedString.ja_JP";
					break;
				case 1031:
					fallBackResourceFile = "msdyn_localizedString.de_DE";
					break;
				case 1036:
					fallBackResourceFile = "msdyn_localizedString.fr_FR";
					break;
				case 1034:
					fallBackResourceFile = "msdyn_localizedString.es_ES";
					break;
				case 1049:
					fallBackResourceFile = "msdyn_localizedString.ru_RU";
					break;
				default:
					fallBackResourceFile = "msdyn_localizedString.en_US";
					break;
				case 1025:
					fallBackResourceFile = "msdyn_localizedString.ar_AR";
					break;
			}
			String ResourceFile = "";
			switch (UserLanguage)
			{
				case 1033:
					ResourceFile = "msdyn_localizedString.en_US";
					break;
				case 1041:
					ResourceFile = "msdyn_localizedString.ja_JP";
					break;
				case 1031:
					ResourceFile = "msdyn_localizedString.de_DE";
					break;
				case 1036:
					ResourceFile = "msdyn_localizedString.fr_FR";
					break;
				case 1034:
					ResourceFile = "msdyn_localizedString.es_ES";
					break;
				case 1049:
					ResourceFile = "msdyn_localizedString.ru_RU";
					break;
				case 1025:
					ResourceFile = "msdyn_localizedString.ar_AR";
					break;
				default:
					ResourceFile = fallBackResourceFile;
					break;
			}

			return ResourceFile;
		}

		/// <summary>
		/// This method returns Organizations base language code.
		/// </summary>
		/// <param name="service"> IOrganizationService object</param>
		/// <returns>Organization language code</returns>
		internal static int RetrieveOrganizationBaseLanguageCode(IOrganizationService service)
		{
			QueryExpression organizationEntityQuery = new QueryExpression("organization");
			organizationEntityQuery.ColumnSet.AddColumn("languagecode");
			EntityCollection organizationEntities = service.RetrieveMultiple(organizationEntityQuery);
			return (int)organizationEntities[0].Attributes["languagecode"];
		}

		/// <summary>
		/// This method returns User UI language code.
		/// </summary>
		/// <param name="service">IOrganizationService object</param>
		/// <param name="userId">User Id</param>
		/// <returns>User languages code</returns>
		internal static int RetrieveUserUILanguageCode(IOrganizationService service, Guid userId)
		{
			QueryExpression userSettingsQuery = new QueryExpression("usersettings");
			userSettingsQuery.ColumnSet.AddColumns("uilanguageid", "systemuserid");
			userSettingsQuery.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
			EntityCollection userSettings = service.RetrieveMultiple(userSettingsQuery);
			if (userSettings.Entities.Count > 0)
			{
				return (int)userSettings.Entities[0]["uilanguageid"];
			}
			return 0;
		}

		/// <summary>
		/// This method extract the XMLDocument element from web resource.
		/// </summary>
		/// <param name="organizationService">IOrganizationService object</param>
		/// <param name="webresourceSchemaName"> Web resource file name</param>
		/// <returns>XML Document element</returns>
		internal static XmlDocument RetrieveXmlWebResourceByName(IOrganizationService organizationService, string webresourceSchemaName)
		{
			QueryExpression webresourceQuery = new QueryExpression("webresource");
			webresourceQuery.ColumnSet.AddColumn("content");
			webresourceQuery.Criteria.AddCondition("name", ConditionOperator.Equal, webresourceSchemaName);
			EntityCollection webresources = organizationService.RetrieveMultiple(webresourceQuery);
			if (webresources.Entities.Count > 0)
			{
				byte[] bytes = Convert.FromBase64String((string)webresources.Entities[0]["content"]);
				// The bytes would contain the ByteOrderMask. Encoding.UTF8.GetString() does not remove the BOM.
				// Stream Reader auto detects the BOM and removes it on the text
				XmlDocument document = new XmlDocument();
				document.XmlResolver = null;
				using (MemoryStream ms = new MemoryStream(bytes))
				{
					using (StreamReader sr = new StreamReader(ms))
					{
						document.Load(sr);
					}
				}
				return document;
			}
			else
			{
				throw new InvalidPluginExecutionException(String.Format("Unable to locate the web resource {0}.", webresourceSchemaName));
			}
		}

		/// <summary>
		/// This method returns inner text of XMLDocument node.
		/// </summary>
		/// <param name="resource">XmlDocument node</param>
		/// <param name="resourceId">resourceId</param>
		/// <returns>Localized string</returns>
		internal static string RetrieveLocalizedStringFromWebResource(XmlDocument resource, string resourceId)
		{
			XmlNode valueNode = resource.SelectSingleNode(string.Format(CultureInfo.InvariantCulture, "./root/data[@name='{0}']/value", resourceId));
			if (valueNode != null)
			{
				return valueNode.InnerText;
			}
			else
			{
				throw new InvalidPluginExecutionException(String.Format("ResourceID {0} was not found.", resourceId));
			}
		}
	}
}