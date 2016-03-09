using System;
using System.Collections.Generic;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace KnowledgeArticleTranslationSolution
{
	/// <summary>
	/// This is the plugin to handle the logic of auto translating the Knowledge Article post creation of translation
	/// </summary>
	public class AutoTranslationPlugin : IPlugin
	{
		public void Execute(IServiceProvider serviceProvider)
		{
			//Extract the tracing service for use in debugging sandboxed plug-ins.
			ITracingService tracingService =
				(ITracingService)serviceProvider.GetService(typeof(ITracingService));

			// Obtain the execution context from the service provider.
			IPluginExecutionContext context = (IPluginExecutionContext)
				serviceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service reference.
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			// The InputParameters collection contains all the data passed in the message request.
			if (context.InputParameters.Contains("Target") &&
				context.InputParameters["Target"] is Entity)
			{
				// Obtain the target entity from the input parameters.
				Entity translateEntity = (Entity)context.InputParameters["Target"];

				TranslateApiSettings apiSettings = MicrosoftTranslateHelper.GetTranslatorAPISettings(service, tracingService);

				// Update the knowledge article entity record if isautotranslationenabled is enabled
				// And it shouldn't be primary and root article.

				if (apiSettings != null &&
					apiSettings.Isautotranslationenabled == true &&
					((bool)translateEntity["isprimary"]) == false &&
					((bool)translateEntity["isrootarticle"]) == false)
				{
					ColumnSet columns = new ColumnSet();
					columns.AddColumn("languagelocaleid");

					// Get the primary entity to retrieve the source language code
					EntityReference parentArticleReference = (EntityReference)translateEntity["parentarticlecontentid"];
					Entity primaryEntity = service.Retrieve("knowledgearticle", parentArticleReference.Id, columns);
					EntityReference languageLocaleReference = (EntityReference)primaryEntity["languagelocaleid"];

					columns = new ColumnSet();
					columns.AddColumn("code");
					Entity languageLocaleEntity = service.Retrieve("languagelocale", languageLocaleReference.Id, columns);
					string sourceLanguageCode = languageLocaleEntity["code"].ToString();

					// Get the language code of the knowledge article being translated
					languageLocaleReference = (EntityReference)translateEntity["languagelocaleid"];
					languageLocaleEntity = service.Retrieve("languagelocale", languageLocaleReference.Id, columns);
					string destinationLanguageCode = languageLocaleEntity["code"].ToString();

					tracingService.Trace("knowledgeArticleId of the parent article: " + parentArticleReference.Id.ToString());
					tracingService.Trace("Translating from <" + sourceLanguageCode + "> to <" + destinationLanguageCode + ">");
					// Start Translation process
					try
					{
						if (translateEntity.Attributes.ContainsKey("title"))
						{
							translateEntity["title"] = MicrosoftTranslateHelper.Translate(translateEntity["title"].ToString(), sourceLanguageCode, destinationLanguageCode, apiSettings.ClientSecret, apiSettings.ClientID);
						}
						tracingService.Trace("Translated <title>: {0}", translateEntity["title"].ToString());

						if (translateEntity.Attributes.ContainsKey("keywords"))
						{
							translateEntity["keywords"] = MicrosoftTranslateHelper.Translate(translateEntity["keywords"].ToString(), sourceLanguageCode, destinationLanguageCode, apiSettings.ClientSecret, apiSettings.ClientID);
						}

						if (translateEntity.Attributes.ContainsKey("description"))
						{
							translateEntity["description"] = MicrosoftTranslateHelper.Translate(translateEntity["description"].ToString(), sourceLanguageCode, destinationLanguageCode, apiSettings.ClientSecret, apiSettings.ClientID);
						}

						if (translateEntity.Attributes.ContainsKey("content"))
						{
							translateEntity["content"] = MicrosoftTranslateHelper.Translate(translateEntity["content"].ToString(), sourceLanguageCode, destinationLanguageCode, apiSettings.ClientSecret, apiSettings.ClientID);
						}
					}
					catch (FaultException<OrganizationServiceFault> ex)
					{
						tracingService.Trace("AutomaticKnowledgeArticleTranslationSolution: {0}", ex.ToString());
						throw new Exception(ex.Message);
					}

					catch (Exception ex)
					{
						tracingService.Trace("AutomaticKnowledgeArticleTranslationSolution: {0}", ex.ToString());
						throw new Exception(ex.Message);
					}
				}
			}
		}
	}
}
