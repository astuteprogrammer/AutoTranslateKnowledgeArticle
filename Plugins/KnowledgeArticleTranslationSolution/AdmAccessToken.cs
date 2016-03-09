using System.Runtime.Serialization;

namespace KnowledgeArticleTranslationSolution
{
	/// <summary>
	/// Translator Access Token Class
	/// </summary>
	[DataContract]
	public class AdmAccessToken
	{
		[DataMember]
		public string access_token { get; set; }
		[DataMember]
		public string token_type { get; set; }
		[DataMember]
		public string expires_in { get; set; }
		[DataMember]
		public string scope { get; set; }
	}
}
