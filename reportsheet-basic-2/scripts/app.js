// 日本語カルチャー設定
GC.Spread.Common.CultureManager.culture("ja-jp");
//GC.Spread.Sheets.LicenseKey = "ここにSpreadJSのライセンスキーを設定します";

// SpreadJSの設定
document.addEventListener("DOMContentLoaded", () => {
    const spread = new GC.Spread.Sheets.Workbook("ss");
    const printButton = document.getElementById('print');
    const excelButton = document.getElementById('excel');
    const pdfButton = document.getElementById('pdf');

    let reportSheet;

    //------------------------------------------
    // PDFエクスポートに必要なフォントを登録します
    //------------------------------------------
    registerFont("IPAexゴシック", "normal", "fonts/ipaexg.ttf");

    //----------------------------------------------------------------
    // sjs形式のテンプレートシートを読み込んでレポートシートを実行します
    //----------------------------------------------------------------
    const res = fetch('reports/products.sjs').then((response) => response.blob())
        .then((myBlob) => {
            spread.open(myBlob, () => {
                console.log(`読み込み成功`);
                reportSheet = spread.getSheetTab(0);

                // レポートシートのオプション設定
                reportSheet.renderMode('PaginatedPreview');
                reportSheet.options.printAllPages = true;

                // レポートシートの印刷設定
                var printInfo = reportSheet.printInfo();
                printInfo.showBorder(false);
                printInfo.zoomFactor(1);
                reportSheet.printInfo(printInfo);
            }, (e) => {
                console.log(`***ERR*** エラーコード（${e.errorCode}） : ${e.errorMessage}`);
            });
        });
    //------------------------------------------
    // 印刷ボタン押下時の処理
    //------------------------------------------
    printButton.onclick = function () {
        spread.print();
    }

    //------------------------------------------
    // Excelエクスポートボタン押下時の処理
    //------------------------------------------
    excelButton.onclick = function () {
        spread.export(function (blob) { saveAs(blob, 'products.xlsx'); }, function (error) {
            console.log(error);
        }, {
            fileType: GC.Spread.Sheets.FileType.excel,
        })
    }
    //------------------------------------------
    // PDF出力ボタン押下時の処理
    //------------------------------------------
    pdfButton.onclick = function () {
        spread.savePDF(function (blob) {
            //saveAs(blob, 'products.pdf');
            const url = URL.createObjectURL(blob);
            window.open(url);
        }, function (error) {
            console.log(error);
        }, {
            title: '商品一覧',
            author: 'Test Author',
            subject: 'Test Subject',
            keywords: 'Test Keywords',
            creator: 'Test Creator'
        })
    }

    //------------------------------------------
    // フォントファイルの読み込み
    //------------------------------------------      
    function registerFont(name, type, fontPath) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', fontPath, true);
        xhr.responseType = 'arraybuffer';
        xhr.onload = function (e) {
            if (this.status == 200) {
                var fontArrayBuffer = this.response;
                var fonts = {};
                fonts[type] = fontArrayBuffer;
                GC.Spread.Sheets.PDF.PDFFontsManager.registerFont(name, fonts);
            }
        };
        xhr.send();
    }

});

