namespace KnowledgeArticleTranslationSolution
{
	/// <summary>
	/// Translate API Settings Class
	/// </summary>
	public class TranslateApiSettings
	{
		string clientID;

		public string ClientID
		{
			get { return clientID; }
			set { clientID = value; }
		}

		string clientSecret;

		public string ClientSecret
		{
			get { return clientSecret; }
			set { clientSecret = value; }
		}

		bool isautotranslationenabled;

		public bool Isautotranslationenabled
		{
			get { return isautotranslationenabled; }
			set { isautotranslationenabled = value; }
		}

	}
}
