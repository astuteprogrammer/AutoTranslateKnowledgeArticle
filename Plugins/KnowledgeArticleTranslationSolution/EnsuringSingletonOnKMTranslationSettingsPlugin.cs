using Microsoft.Xrm.Sdk;
using System;
using System.Xml;

namespace KnowledgeArticleTranslationSolution
{
	/// <summary>
	/// This is the plugin to make sure to have singleton object for Translation settings entity.
	/// </summary>
	public class EnsuringSingletonOnKMTranslationSettingsPlugin : IPlugin
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

				if (translateEntity.LogicalName == "msdyn_automatickmtranslationsetting")
				{
					if (MicrosoftTranslateHelper.GetTranslatorAPISettingsRecordsCount(service, tracingService) == 1)
					{
						string resourceFile = MicrosoftTranslateHelper.GetLanguageResourceFile(service, context);
						XmlDocument messages = MicrosoftTranslateHelper.RetrieveXmlWebResourceByName(service, resourceFile);
						String message = MicrosoftTranslateHelper.RetrieveLocalizedStringFromWebResource(messages, "ErrorMessage");
						throw new Exception(message);
					}
				}
			}
		}
	}
}
