using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace TestesDynamics
{
    public class PreDelete_NotaFiscal : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {
                EntityReference entity = (EntityReference)context.InputParameters["Target"];

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    var notaFiscalId = ((EntityReference)context.InputParameters["Target"]).Id;

                    // Consulta para buscar os registros de custom_msisdn relacionados ao notaFiscalId
                    QueryExpression queryMsisdn = new QueryExpression("custom_msisdn")
                    {
                        ColumnSet = new ColumnSet(false),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("custom_notafiscal", ConditionOperator.Equal, notaFiscalId)
                            }
                        }
                    };

                    // Executa a consulta
                    EntityCollection msisdnRecords = service.RetrieveMultiple(queryMsisdn);

                    // Verifica a quantidade de registros
                    if (msisdnRecords.Entities.Count > 0)
                    {
                        throw new InvalidPluginExecutionException("Impossível excluir essa Nota Fiscal, pois há registros de MSISDN relacionados.");
                    }

                }

                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException($"Erro em PreDelete_NotaFiscal: {ex.Message}");
                }
            }
        }
    }
}
