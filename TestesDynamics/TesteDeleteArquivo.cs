using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TestesDynamics
{
    public class TesteDeleteArquivo : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Valida os parâmetros de entrada
                if (!context.InputParameters.Contains("notaFiscalId") || !(context.InputParameters["notaFiscalId"] is Guid))
                {
                    throw new InvalidPluginExecutionException("Parâmetro 'notaFiscalId' é obrigatório e deve ser do tipo Guid.");
                }

                Guid notaFiscalId = (Guid)context.InputParameters["notaFiscalId"];

                // Consulta para buscar os registros de custom_msisdn relacionados ao notaFiscalId
                QueryExpression queryMsisdn = new QueryExpression("custom_msisdn")
                {
                    ColumnSet = new ColumnSet("custom_msisdnid", "custom_notafiscal"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("custom_notafiscal", ConditionOperator.Equal, notaFiscalId)
                        }
                    }
                };

                // Executa a consulta e obtém os registros
                EntityCollection msisdnRecords = service.RetrieveMultiple(queryMsisdn);
                if (msisdnRecords.Entities.Count == 0)
                {
                    throw new InvalidPluginExecutionException("Erro de Deletar Romaneio: Nenhum registro de MSISDN encontrado relacionada a Nota Fiscal.");
                }

                // Consulta para buscar os registros de custom_depositomsisdn relacionados ao notaFiscalId
                QueryExpression queryDeposito = new QueryExpression("custom_depositomsisdn")
                {
                    ColumnSet = new ColumnSet("custom_depositomsisdnid", "custom_notafiscal"), // Substitua pelos campos que você deseja trazer
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("custom_notafiscal", ConditionOperator.Equal, notaFiscalId)
                        }
                    }
                };

                // Executa a consulta e obtém os registros
                EntityCollection depositoRecords = service.RetrieveMultiple(queryDeposito);
                if (depositoRecords.Entities.Count == 0)
                {
                    throw new InvalidPluginExecutionException("Erro ao Deletar Romaneio: Nenhum registro de Depósito MSISDN encontrado relacionada a Nota Fiscal.");
                }

                // Começa a exclusão dos registros de MSISDN
                List<Guid> msisdnIdsToDelete = new List<Guid>();
                List<Entity> depositoMsisdnToUpdate = new List<Entity>();
                foreach (var msisdn in msisdnRecords.Entities)
                {
                    msisdnIdsToDelete.Add(msisdn.Id);
                }

                // Atualiza os registros de Depósito MSISDN
                foreach (var deposito in depositoRecords.Entities)
                {
                    Entity updateEntity = new Entity("custom_depositomsisdn", deposito.Id);
                    updateEntity["statuscode"] = new OptionSetValue(1);
                    updateEntity["custom_notafiscal"] = null;
                    depositoMsisdnToUpdate.Add(updateEntity);
                }

                // Apaga os registros de MSISDN
                foreach (var msisdnId in msisdnIdsToDelete)
                {
                    service.Delete("custom_msisdn", msisdnId);
                }

                //Atualiza os registros de Depósito MSISDN
                foreach (var depositoMsisdn in depositoMsisdnToUpdate)
                {
                    service.Update(depositoMsisdn);
                }

                //Retorno da api
                context.OutputParameters["resultMessage"] = "Exclusão de dados do romaneio realizada com sucesso!";
            }
            
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
