var metabase, prxReport, prxMbService, dataArea;

function renderReport() {
    const urlSearchParams = new URLSearchParams(window.location.search);
    const params = Object.fromEntries(urlSearchParams.entries());
    if (params.OBJ || params.ID) {
        PP.ImagePath = "http://sng-forweb-dev:8109/fp9.x/build/img/";
        PP.ScriptPath = "http://sng-forweb-dev:8109/fp9.x/build/"; // путь до сборок
        PP.CSSPath = "http://sng-forweb-dev:8109/fp9.x/build/";
        PP.resourceManager.setRootResourcesFolder(
            "http://sng-forweb-dev:8109/fp9.x/resources/"); //путь до папки с ресурсами
        PP.resourceManager.setResourceList(['PP', 'Metabase', 'Regular', 'VisualizerMaster']);
        PP.setCurrentCulture(PP.Cultures.ru); //языковые настройки для ресурсов
        var waiter = new PP.Ui.Waiter();
        metabase = new PP.Mb.Metabase( //создаем соединение с метабазой
            {
                PPServiceUrl: "./service.jsp",
                Id: "SNG_METABASE",
                UserCreds: {
                    UserName: "TEST",
                    Password: "TEST"
                },
                //в начале запроса к метабазе отображается компонент Waiter
                StartRequest: function () {
                    waiter.show();
                },
                //при окончании запроса к метабазе компонент Waiter скрывается
                EndRequest: function (sender, args) {
                    waiter.hide();
                    if (args.Response && args.Response.OpenMetabaseResult) {
                        localStorage.setItem('fs-moniker', JSON.stringify({
                            id: metabase.getConnectionOdId().id,
                            time: new Date().getTime()
                        }))
                        console.log("НОВАЯ СЕССИЯ: " + metabase.getConnectionOdId().id + " от " +
                            new Date());
                    }
                    //при окончании выполнения запроса все запросы удаляются из кэша
                    metabase.clearCache();
                },
                //при ошибке на экране появится сообщение с текстом ошибки
                Error: function (sender, args) {
                    alert(args.ResponseText);
                }
            });
        // Сохраняю моникер в localStorage
        if (localStorage.getItem('fs-moniker')) {
            let oObj = null;
            try {
                let oTemp = JSON.parse(localStorage.getItem('fs-moniker'));
                if ((new Date().getTime() - oTemp.time) <= 500000) oObj = oTemp;
            } catch (e) { }
            if (oObj) {
                metabase._ConnectionOdId.id = oObj.id;
                metabase.refresh();
                console.log("СТАРАЯ СЕССИЯ: " + oObj.id + " от " + new Date(oObj.time));
            } else {
                metabase.open();
            }
        } else {
            metabase.open();
        }


        //открываем метабазу
        metabase.findObjects(113427, {}, function (sender, args) {
            var loadedObjects = JSON.parse(args.ResponseText).GetObjectsResult.mds.its.it;
            let oPrx;
            loadedObjects.find(el => {
                const {
                    obInst,
                    pars
                } = el;
                const {
                    c: iClassId,
                    i: sObjId,
                    k: iId
                } = obInst.obDesc;
                let bOut = iClassId === 2562 &&
                    (!params.OBJ || (params.OBJ && sObjId == params.OBJ)) &&
                    (!params.ID || (params.ID && iId == params.ID));
                if (bOut) {
                    let aParams = [];
                    if (!!pars) {
                        let getValue = ({ id, type }) => {
                            switch(type) {
                                case 2:
                                    return +params[id];
                                default:
                                    return params[id];
                            }
                        };
                        aParams = pars.it.map(({
                            binding,
                            dt,
                            id,
                            k,
                            n,
                            vis
                        }) => {
                            return new PP.Mb.Param({
                                Id: id, // Идентификатор параметра
                                Key: k, // Ключ
                                Name: n, // Наименование параметра
                                Type: dt,
                                Value: getValue({ id, type: dt }),
                                Visible: vis
                            })
                        })
                    }
                    oPrx = {
                        iClassId,
                        sObjId,
                        iId,
                        aParams
                    }
                }
            });
            console.log(oPrx);
            if (oPrx) {
                prxMbService = new PP.Prx.PrxMdService({
                    Metabase: metabase
                }); //создаем сервис для работы с регламентными отчетами
                prxReport = prxMbService.openReport(oPrx.iId, oPrx.aParams, false); //открываем отчет из метабазы по ключу
                //получаем массив параметров
                dataArea = new PP.Prx.Ui.DataArea({
                    ParentNode: "dataArea",
                    Source: prxReport, //указываем отчет-источник
                    Service: prxMbService
                });
                window.onresize();
            } else {
                alert("Отсутствует отчет с такими параметрами")
            }
        })

        function exportXls() {
            let exportData = {
                "format": "xls",
                "storeResult": true,
                "reportTitle": false,
                "sheetTitles": false,
                "showWarnings": true,
                "pagesKeysRange": "",
                "hiddenCells": true,
                "hiddenSheets": false,
                "chartsAsImages": false,
                "sheetAccess": false,
                "formulas": true,
                "exportExpanders": false
            };
            prxMbService.Export(prxReport, exportData, res => {
                try {
                    let oResponse = JSON.parse(res._ResponseText);
                    let oPrxMdResult = !!oResponse.GetPrxMdResult ? oResponse.GetPrxMdResult : oResponse
                        .BatchExecResult.its.it[1].GetPrxMdResult
                    let sId = oPrxMdResult.meta.exportData.storeId.id;
                    //let url = `./serviceExcel.jsp?id=${sId}`;
                    let url = `http://sng-forweb-dev:8109/fp9.x/app/PPService.axd/GetBin?mon=${sId}&fileName=13.7.xls&attach=1`;
                    const anchor = document.createElement('a');
                    anchor.href = url;
                    anchor.download = "TEST.xls";

                    // Append to the DOM
                    document.body.appendChild(anchor);

                    // Trigger `click` event
                    anchor.click();

                    // Remove element from DOM
                    document.body.removeChild(anchor);
                    console.log(sId);
                    // TODO: Метод возвращает ID документа, который нужно передать в PPService.axd
                    // 

                } catch (e) {
                    console.error(e);
                }
            })
        }

        function exportPdf() {
            let exportData = {
                "format": "pdf",
                "storeResult": true,
                "reportTitle": false,
                "sheetTitles": false,
                "showWarnings": true,
                "pagesKeysRange": ""
            };
            prxMbService.Export(prxReport, exportData, res => {
                try {
                    let oResponse = JSON.parse(res._ResponseText);
                    let oPrxMdResult = !!oResponse.GetPrxMdResult ? oResponse.GetPrxMdResult : oResponse
                        .BatchExecResult.its.it[1].GetPrxMdResult
                    let sId = oPrxMdResult.meta.exportData.storeId.id;
                    //let url = `./serviceExcel.jsp?id=${sId}`;
                    let url = `http://sng-forweb-dev:8109/fp9.x/app/PPService.axd/GetBin?mon=${sId}&fileName=13.7.pdf&attach=1`;
                    const anchor = document.createElement('a');
                    anchor.href = url;
                    anchor.download = "TEST.pdf";

                    // Append to the DOM
                    document.body.appendChild(anchor);

                    // Trigger `click` event
                    anchor.click();

                    // Remove element from DOM
                    document.body.removeChild(anchor);
                    console.log(sId);
                    // TODO: Метод возвращает ID документа, который нужно передать в PPService.axd
                    //
                } catch (e) {
                    console.error(e);
                }
            })
        }
        var btnXls = new PP.Ui.Button({
            ParentNode: document.getElementById("btn1"),
            Click: exportXls,
            Content: "Экспорт в Excel"
        })
        var btnPdf = new PP.Ui.Button({
            ParentNode: document.getElementById("btn2"),
            Click: exportPdf,
            Content: "Экспорт в PDF"
        })
        window.onresize();
    } else {
        alert("Отсутствует обязательный параметр: ID или OBJ");
    }
}
var idTime;
//функция для изменения размера компонента при изменении размера контейнера
window.onresize = function updateSize() {
    if (idTime)
        clearTimeout(idTime);
    idTime = setTimeout(function () {
        if (dataArea) {
            let oDataArea = document.getElementById("dataArea");
            oDataArea.style["height"] = `${window.innerHeight - 23}px`;
            dataArea.setSize(oDataArea.offsetWidth - 2, window.innerHeight - 23);
        }
        idTime = null;
    }, 50);
};