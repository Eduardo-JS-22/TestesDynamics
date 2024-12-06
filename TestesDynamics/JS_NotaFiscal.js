var NotaFiscal = {
    Events: {
        BotaoUploadRomaneio: function (context) {
            var entityReference = context.data.entity.getEntityReference();
            var notaFiscalId = entityReference.id.replace("{", "").replace("}", "");
            Xrm.Navigation.openAlertDialog({ text: "Upload de Romaneio em Curso", title: "Upload de Romaneio" });
            NotaFiscal.Methods.ChamarApiUploadRomaneio(notaFiscalId);
        },

        BotaoDeleteRomaneio: function (context) {
            var entityReference = context.data.entity.getEntityReference();
            var notaFiscalId = entityReference.id.replace("{", "").replace("}", "");
            Xrm.Navigation.openAlertDialog({ text: "Delete de Romaneio em Curso", title: "Delete de Romaneio" });
            NotaFiscal.Methods.ChamarApiDeleteRomaneio(notaFiscalId);
        },

        BotaoUploadNotaFiscal: function (context) {
            var entityReference = context.data.entity.getEntityReference();
            var notaFiscalId = entityReference.id.replace("{", "").replace("}", "");
            Xrm.Navigation.openAlertDialog({ text: "Upload de Nota Fiscal em Curso", title: "Upload de Nota Fiscal" });
            NotaFiscal.Methods.ChamarApiUploadNotaFiscal(notaFiscalId);
        },

        BotaoDeleteNotaFiscal: function (context) {
            var entityReference = context.data.entity.getEntityReference();
            var notaFiscalId = entityReference.id.replace("{", "").replace("}", "");
            Xrm.Navigation.openAlertDialog({ text: "Delete de Nota Fiscal em Curso", title: "Delete de Nota Fiscal" });
            NotaFiscal.Methods.ChamarApiDeleteNotaFiscal(notaFiscalId);
        }
    },

    Methods: {
        ChamarApiUploadRomaneio: function (notaFiscalId) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/custom_UploadRomaneioApi(notaFiscalId=@notaFiscalId)?@notaFiscalId=" + notaFiscalId, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.response);
                        console.log(result);
                        var resultmessage = result["resultMessage"];
                        Xrm.Navigation.openAlertDialog({ text: resultmessage, title: "Upload de Romaneio" });
                    } else {
                        console.log(this.responseText);
                        Xrm.Navigation.openErrorDialog({ message: this.responseText });
                    }
                }
            };
            req.send();
        },

        ChamarApiDeleteRomaneio: function (notaFiscalId) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/custom_DeleteUploadRomaneio(notaFiscalId=@notaFiscalId)?@notaFiscalId=" + notaFiscalId, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.response);
                        console.log(result);
                        var resultmessage = result["resultMessage"];
                        Xrm.Navigation.openAlertDialog({ text: resultmessage, title: "Delete de Romaneio" });
                    } else {
                        console.log(this.responseText);
                        Xrm.Navigation.openErrorDialog({ message: this.responseText });
                    }
                }
            };
            req.send();
        },

        ChamarApiUploadNotaFiscal: function (notaFiscalId) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/custom_UploadNotaFiscal(notaFiscalId=@notaFiscalId)?@notaFiscalId=" + notaFiscalId, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.response);
                        console.log(result);
                        var resultmessage = result["resultMessage"];
                        Xrm.Navigation.openAlertDialog({ text: resultmessage, title: "Upload de Nota Fiscal" });
                    } else {
                        console.log(this.responseText);
                        Xrm.Navigation.openErrorDialog({ message: this.responseText });
                    }
                }
            };
            req.send();
        },

        ChamarApiDeleteNotaFiscal: function (notaFiscalId) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/custom_DeleteNotaFiscal(notaFiscalId=@notaFiscalId)?@notaFiscalId=" + notaFiscalId, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.response);
                        console.log(result);
                        var resultmessage = result["resultMessage"];
                        Xrm.Navigation.openAlertDialog({ text: resultmessage, title: "Delete de Nota Fiscal" });
                    } else {
                        console.log(this.responseText);
                        Xrm.Navigation.openErrorDialog({ message: this.responseText });
                    }
                }
            };
            req.send();
        }
    }
}