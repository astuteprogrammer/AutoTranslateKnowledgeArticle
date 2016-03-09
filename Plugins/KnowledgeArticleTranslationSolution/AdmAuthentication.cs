using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;

namespace KnowledgeArticleTranslationSolution
{
	/// <summary>
	/// Translator Authentication Class
	/// </summary>
	public class AdmAuthentication
	{

		public static readonly string DatamarketAccessUri = "https://datamarket.accesscontrol.windows.net/v2/OAuth2-13";
		private string clientId;
		private string clientSecret;
		private string request;

		/// <summary>
		/// Authenticate with the Microsoft Translate.
		/// </summary>
		/// <param name="clientId">Client ID</param>
		/// <param name="clientSecret">Client Secret</param>
		public AdmAuthentication(string clientId, string clientSecret)
		{
			this.clientId = clientId;
			this.clientSecret = clientSecret;
			//If clientid or client secret has special characters, encode before sending request
			this.request = string.Format("grant_type=client_credentials&client_id={0}&client_secret={1}&scope=http://api.microsofttranslator.com", System.Net.WebUtility.UrlEncode(clientId), System.Net.WebUtility.UrlEncode(clientSecret));
		}

		/// <summary>
		/// Returns the access token
		/// </summary>
		/// <returns>AdmAccessToken object</returns>
		public AdmAccessToken GetAccessToken()
		{
			return HttpPost(DatamarketAccessUri, this.request);
		}

		/// <summary>
		/// This mehod used to post the request.
		/// </summary>
		/// <param name="DatamarketAccessUri">DatamarketAccessUri</param>
		/// <param name="requestDetails">requestDetails</param>
		/// <returns>AdmAccessToken object</returns>
		private AdmAccessToken HttpPost(string DatamarketAccessUri, string requestDetails)
		{
			//Prepare OAuth request 
			WebRequest webRequest = WebRequest.Create(DatamarketAccessUri);
			webRequest.ContentType = "application/x-www-form-urlencoded";
			webRequest.Method = "POST";
			byte[] bytes = Encoding.ASCII.GetBytes(requestDetails);
			webRequest.ContentLength = bytes.Length;
			using (Stream outputStream = webRequest.GetRequestStream())
			{
				outputStream.Write(bytes, 0, bytes.Length);
			}
			using (WebResponse webResponse = webRequest.GetResponse())
			{
				DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AdmAccessToken));
				//Get deserialized object from JSON stream
				AdmAccessToken token = (AdmAccessToken)serializer.ReadObject(webResponse.GetResponseStream());
				return token;
			}
		}
	}
}
