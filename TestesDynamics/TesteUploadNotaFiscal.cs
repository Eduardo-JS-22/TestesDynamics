using System;
using System.Collections.Generic;
using System.Xml;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TestesDynamics
{
    public class TesteUploadNotaFIscal : IPlugin
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
                Entity notaFiscal = service.Retrieve("custom_notafiscal", notaFiscalId, new ColumnSet("custom_arquivo_nota_fiscal"));
                if (notaFiscal == null || !notaFiscal.Contains("custom_arquivo_nota_fiscal"))
                {
                    throw new InvalidPluginExecutionException("Arquivo Nota Fiscal não encontrado para a Nota Fiscal fornecida.");
                }

                // Faz o download do conteúdo do campo de arquivo
                var fileContent = DownloadFile(service, new EntityReference("custom_notafiscal", notaFiscalId), "custom_arquivo_nota_fiscal");
                string xmlContent = System.Text.Encoding.UTF8.GetString(fileContent);

                // Carrega o conteúdo do arquivo como XML
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlContent);

                // Busca o nó <dest>
                XmlNode destNode = xmlDocument.SelectSingleNode("//dest");
                if (destNode == null)
                {
                    throw new InvalidPluginExecutionException("Nó <dest> não encontrado no XML.");
                }

                // Extrai os dados do nó <dest>
                string nome = destNode.SelectSingleNode("nome")?.InnerText;
                string documento = destNode.SelectSingleNode("documento")?.InnerText;
                string telefone = destNode.SelectSingleNode("telefone")?.InnerText;
                string email = destNode.SelectSingleNode("email")?.InnerText;
                string endereco = destNode.SelectSingleNode("endereco")?.InnerText;

                // Atualiza os campos na Nota Fiscal
                Entity notaFiscalUpdate = new Entity("custom_notafiscal", notaFiscalId);
                if (!string.IsNullOrWhiteSpace(nome)) notaFiscalUpdate["custom_nome"] = nome;
                if (!string.IsNullOrWhiteSpace(documento)) notaFiscalUpdate["custom_documento"] = documento;
                if (!string.IsNullOrWhiteSpace(telefone)) notaFiscalUpdate["custom_telefone"] = telefone;
                if (!string.IsNullOrWhiteSpace(email)) notaFiscalUpdate["custom_email"] = email;
                if (!string.IsNullOrWhiteSpace(endereco)) notaFiscalUpdate["custom_endereco"] = endereco;

                service.Update(notaFiscalUpdate);

                //Retorno da api
                context.OutputParameters["resultMessage"] = "Upload de dados da Nota Fiscal realizada com sucesso!";
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Erro no Upload de Nota Fiscal: {ex.Message}");
            }
        }

        private static byte[] DownloadFile(IOrganizationService service, EntityReference entityReference, string attributeName)
        {
            try
            {
                InitializeFileBlocksDownloadRequest initializeFileBlocksDownloadRequest = new InitializeFileBlocksDownloadRequest
                {
                    Target = entityReference,
                    FileAttributeName = attributeName
                };

                var initializeFileBlocksDownloadResponse = (InitializeFileBlocksDownloadResponse)service.Execute(initializeFileBlocksDownloadRequest);

                string fileContinuationToken = initializeFileBlocksDownloadResponse.FileContinuationToken;
                long fileSizeInBytes = initializeFileBlocksDownloadResponse.FileSizeInBytes;

                List<byte> fileBytes = new List<byte>((int)fileSizeInBytes);

                long offset = 0;
                // Se o chunking não for suportado, o tamanho do bloco será igual ao tamanho total do arquivo.
                long blockSizeDownload = !initializeFileBlocksDownloadResponse.IsChunkingSupported ? fileSizeInBytes : 4 * 1024 * 1024;

                if (fileSizeInBytes < blockSizeDownload)
                {
                    blockSizeDownload = fileSizeInBytes;
                }

                while (fileSizeInBytes > 0)
                {
                    // Preparar a solicitação
                    DownloadBlockRequest downLoadBlockRequest = new DownloadBlockRequest
                    {
                        BlockLength = blockSizeDownload,
                        FileContinuationToken = fileContinuationToken,
                        Offset = offset
                    };

                    // Enviar a solicitação
                    var downloadBlockResponse = (DownloadBlockResponse)service.Execute(downLoadBlockRequest);

                    // Adicionar o bloco retornado à lista
                    fileBytes.AddRange(downloadBlockResponse.Data);

                    // Subtrair a quantidade de bytes baixados
                    fileSizeInBytes -= (int)blockSizeDownload;
                    // Incrementar o offset para começar do início do próximo bloco.
                    offset += blockSizeDownload;
                }

                return fileBytes.ToArray();
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Erro no plugin de Upload de Romaneio", ex);
            }
        }
    }
}
