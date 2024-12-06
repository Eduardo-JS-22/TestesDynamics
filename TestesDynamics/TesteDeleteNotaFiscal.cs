using System;
using System.Collections.Generic;
using System.Xml;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TestesDynamics
{
    public class TesteDeleteNotaFIscal : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Valida o parâmetro de entrada (notaFiscalId)
                if (!context.InputParameters.Contains("notaFiscalId") || !(context.InputParameters["notaFiscalId"] is Guid))
                {
                    throw new InvalidPluginExecutionException("Parâmetro 'notaFiscalId' é obrigatório e deve ser do tipo Guid.");
                }

                Guid notaFiscalId = (Guid)context.InputParameters["notaFiscalId"];

                // Busca a Nota Fiscal e verifica se o campo de arquivo existe
                Entity notaFiscal = service.Retrieve("custom_notafiscal", notaFiscalId, new ColumnSet(false));
                if (notaFiscal == null)
                {
                    throw new InvalidPluginExecutionException("Nota Fiscal não encontrado para a Nota Fiscal fornecida.");
                }

                // Atualiza os campos na Nota Fiscal
                Entity notaFiscalUpdate = new Entity("custom_notafiscal", notaFiscalId);
                notaFiscalUpdate["custom_nome"] = "";
                notaFiscalUpdate["custom_documento"] = "";
                notaFiscalUpdate["custom_telefone"] = "";
                notaFiscalUpdate["custom_email"] = "";
                notaFiscalUpdate["custom_endereco"] = "";
                service.Update(notaFiscalUpdate);

                //Retorno da api
                context.OutputParameters["resultMessage"] = "Delete de dados da Nota Fiscal realizada com sucesso!";
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Erro no Delete de Nota Fiscal: {ex.Message}");
            }
        }
    }
}
