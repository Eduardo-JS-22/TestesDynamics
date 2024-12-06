using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text;

namespace TestesDynamics
{
    public class TesteArquivoApi : IPlugin
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

                // Recupera a nota fiscal
                Entity notaFiscal = service.Retrieve("custom_notafiscal", notaFiscalId, new ColumnSet("custom_arquivo_romaneio", "custom_notafiscalid", "custom_ordem"));

                if (notaFiscal.Contains("custom_arquivo_romaneio"))
                {
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
                    if (msisdnRecords.Entities.Count > 0)
                    {
                        throw new InvalidPluginExecutionException("Erro de Upload de Romaneio: Essa Nota Fiscal já conta com Romaneio.");
                    }

                    Guid arquivoId = notaFiscal.GetAttributeValue<Guid>("custom_arquivo_romaneio");
                    Guid ordemId = notaFiscal.GetAttributeValue<EntityReference>("custom_ordem").Id;

                    // Recupera o conteúdo do arquivo
                    byte[] fileBytes = DownloadFile(service, new EntityReference("custom_notafiscal", notaFiscalId), "custom_arquivo_romaneio");

                    // Processa o conteúdo do arquivo
                    var iccids = ReturnIccidFromExcel(fileBytes);
                    var imeis = ReturnImeiFromExcel(fileBytes);

                    var result = VerificaAtualizaMsisdn(iccids, imeis, notaFiscalId, ordemId, service);

                    if (result != "")
                    {
                        throw new PluginExecutionExceptionApi($"Erro no Upload de Romaneio:\n{result}", "ERR0001");
                    }

                    // Se tudo ocorrer bem, retorna sucesso
                    context.OutputParameters["resultMessage"] = "Upload de romaneio realizado com sucesso!";
                }
            }

            catch (InvalidPluginExecutionException ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

            catch (PluginExecutionException ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
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


        private string[] ReturnIccidFromExcel(byte[] fileBytes)
        {
            try
            {
                // Inicializa a lista para armazenar os valores
                var iccidList = new List<string>();

                // Abre o arquivo Excel a partir do array de bytes
                using (var memoryStream = new MemoryStream(fileBytes))
                {
                    // Abre o arquivo Excel no formato OpenXml
                    using (var document = SpreadsheetDocument.Open(memoryStream, false))
                    {
                        // Obtém a primeira planilha
                        var workbookPart = document.WorkbookPart;
                        var sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                        // Obtém as linhas da planilha
                        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // Itera pelas linhas (a partir da segunda, ignorando o cabeçalho)
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            // Pula a primeira linha (cabeçalho)
                            if (row.RowIndex.Value == 1) continue;

                            // Obtém o valor da segunda célula (coluna do ICCID)
                            var iccidCell = row.Elements<Cell>().ElementAtOrDefault(1); // Coluna 2 (index 1)

                            if (iccidCell != null)
                            {
                                // Obtém o valor da célula, considerando que ela pode ser um valor compartilhado
                                var iccidValue = GetCellValue(document, iccidCell);
                                iccidList.Add(iccidValue);
                            }
                        }
                    }
                }

                // Converte a lista para array
                return iccidList.ToArray();
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Erro no plugin de Upload de Romaneio", ex);
            }
        }

        private string[] ReturnImeiFromExcel(byte[] fileBytes)
        {
            try
            {
                // Inicializa a lista para armazenar os valores
                var imeiList = new List<string>();

                // Abre o arquivo Excel a partir do array de bytes
                using (var memoryStream = new MemoryStream(fileBytes))
                {
                    // Abre o arquivo Excel no formato OpenXml
                    using (var document = SpreadsheetDocument.Open(memoryStream, false))
                    {
                        // Obtém a primeira planilha
                        var workbookPart = document.WorkbookPart;
                        var sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                        // Obtém as linhas da planilha
                        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // Itera pelas linhas (a partir da segunda, ignorando o cabeçalho)
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            // Pula a primeira linha (cabeçalho)
                            if (row.RowIndex.Value == 1) continue;

                            // Obtém o valor da terceira célula (coluna do IMEI)
                            var imeiCell = row.Elements<Cell>()
                                .ElementAtOrDefault(2); // Coluna 3 (index 2)

                            if (imeiCell != null)
                            {
                                // Obtém o valor da célula, considerando que ela pode ser um valor compartilhado
                                var imeiValue = GetCellValue(document, imeiCell);
                                imeiList.Add(imeiValue);
                            }
                        }
                    }
                }

                // Converte a lista para array
                return imeiList.ToArray();
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Erro no plugin de Upload de Romaneio", ex);
            }
        }

        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            try
            {
                // Se a célula tiver um valor compartilhado, encontra o valor
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    var index = int.Parse(cell.CellValue.Text);
                    return sharedStringTable.ElementAt(index).InnerText;
                }

                // Caso contrário, é um valor direto da célula
                return cell.CellValue?.Text;
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Erro no plugin de Upload de Romaneio", ex);
            }
        }

        public string VerificaAtualizaMsisdn(string[] iccids, string[] imeis, Guid notaFiscalId, Guid ordemId, IOrganizationService service)
        {
            try
            {
                // Verifica se as listas de ICCIDs e IMEIs têm o mesmo tamanho
                if (iccids.Length != imeis.Length)
                {
                    return "Erro de Layout";
                }

                // Inicializa listas para atualizações e validações
                var logProcessos = new StringBuilder();
                var registrosParaAtualizar = new List<Entity>();
                var registrosParaCriar = new List<Entity>();
                bool houveErro = false;

                for (int i = 0; i < iccids.Length; i++)
                {
                    string iccid = iccids[i];
                    string imei = imeis[i];

                    // Busca o ICCID na tabela Deposito MSISDN
                    QueryExpression query = new QueryExpression("custom_depositomsisdn")
                    {
                        ColumnSet = new ColumnSet("custom_msisdn", "custom_iccid", "statuscode"),
                        Criteria = new FilterExpression
                        {
                            Conditions = { new ConditionExpression("custom_iccid", ConditionOperator.Equal, iccid) }
                        }
                    };
                    EntityCollection result = service.RetrieveMultiple(query);

                    if (result.Entities.Count == 0)
                    {
                        // Registro ICCID não encontrado
                        logProcessos.AppendLine($"ICCID {iccid} Inexistente\n");
                        houveErro = true;
                        continue;
                    }

                    Entity depositoMsisdn = result.Entities.First();

                    // Verifica se o ICCID está em uso
                    var statusCode = depositoMsisdn.GetAttributeValue<OptionSetValue>("statuscode");
                    if (statusCode == null || statusCode.Value != 1)
                    {
                        logProcessos.AppendLine($"ICCID {iccid} Em Uso");
                        houveErro = true;
                        continue;
                    }

                    // Adiciona registro para atualização na tabela Deposito MSISDN
                    depositoMsisdn["statuscode"] = new OptionSetValue(483050001); // Atualiza para status "Reservado"
                    depositoMsisdn["custom_notafiscal"] = new EntityReference("custom_notafiscal", notaFiscalId);
                    registrosParaAtualizar.Add(depositoMsisdn);

                    // Cria novo registro na tabela MSISDN
                    var novoMsisdn = new Entity("custom_msisdn")
                    {
                        ["custom_msisdn"] = depositoMsisdn.GetAttributeValue<string>("custom_msisdn"),
                        ["custom_imei"] = imei,
                        ["custom_ordem"] = new EntityReference("custom_ordem", ordemId),
                        ["custom_notafiscal"] = new EntityReference("custom_notafiscal", notaFiscalId)
                    };
                    registrosParaCriar.Add(novoMsisdn);
                }

                // Se houve erro, não realiza as atualizações/criações
                if (houveErro)
                {
                    return logProcessos.ToString();
                }

                // Atualiza os registros no Deposito MSISDN
                foreach (var registro in registrosParaAtualizar)
                {
                    service.Update(registro);
                }

                // Cria os registros na tabela MSISDN
                foreach (var registro in registrosParaCriar)
                {
                    service.Create(registro);
                }

                return "";
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Erro no plugin de Upload de Romaneio", ex);
            }
        }
    }

    public class PluginExecutionExceptionApi : Exception
    {
        public string ErrorCode { get; }

        public PluginExecutionExceptionApi(string message, string errorCode) : base(message)
        {
            ErrorCode = errorCode;
        }
    }
}
