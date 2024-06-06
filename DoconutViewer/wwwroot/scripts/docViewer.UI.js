Dropzone.autoDiscover = false;

var objctlDoc = null;
var globalToken = "";
var loader = $(".loader");
var isMobile = false;
var resizing = false;
var docViewerDiv = $("#divDocViewer");
var w = 0;
var h = 0;

loader.hide();

if (/Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)) {
    isMobile = true;
}

function ctlDoc_OnViewerBusy() {
    loader.show();
}

function ctlDoc_OnViewerReady() {
    loader.hide();
}

function RotateDocument(iAngle) {
    objctlDoc.Rotate(objctlDoc.CurrentPage(), iAngle);
}

function FlipDocument(flipType) {
    objctlDoc.Flip(objctlDoc.CurrentPage(), flipType);
}

function Resize() {

    if (resizing) { return; }

    resizing = true;

    var yPadding = 20;

    if (isMobile) {
        yPadding = 30; // change as required
    }

    docViewerDiv.css("height", "90vh");
    docViewerDiv.height(docViewerDiv.height() - yPadding);

    Refit();

    resizing = false;
}

function Refit() {
    if (null !== objctlDoc) {
        objctlDoc.Refit();

        // optional call
        setTimeout(function () { objctlDoc.FitType('width'); }, 1000); // wait 1 seconds
    }
}


function OpenDocument(fileName) {

    loader.show();

    $.ajax({
        type: "POST",
        cache: false,
        async: true,
        url: "/Home/OpenDocument?fileName=" + fileName,
        success: function (data) {

            globalToken = data; // for print
            objctlDoc.View(data);

        },
        error: function (err) {
            alert("Unable to open document. Error: " + err.responseText);
            loader.hide();
        }
    });
}

function CloseDocument() {

    if (globalToken.length == 0) {
        return;
    }

    loader.show();

    $.ajax({
        type: "POST",
        cache: false,
        async: true, 
        url: "/Home/CloseDocument?token=" + globalToken,
        success: function (data) {
            objctlDoc.Close();

            globalToken = "";
            loader.hide();
        },
        error: function (err) {
            alert("Unable to close document. Error: " + err.responseText);
            loader.hide();
        }
    });
}


$(window).on("load", function () {
    objctlDoc = $("#div_ctlDoc").docViewer(
        {
            showThumbs: true,
            autoFocus: true,
            autoPageFocus: false,
            pageZoom: 75,
            zoomStep: 10,
            maxZoom: 200,
            format: 'png',
            FitType: 'width',
            debugMode: true,
            showToolTip: true,
            cacheEnabled: false,
            autoLoad: false,
            toolTipPageText: 'Page ',
            BasePath: '/',
            ResPath: 'images',
            largeDoc: true,
            showHyperlinks: true,
            fixedZoom: true,
            fixedZoomPercent: 100,
            fixedZoomPercentMobile: 75
        });

    Resize();
});

$(window).on("resize orientationchange", function () { Resize(); });

function PrintDocument() {
    if (globalToken.length == 0) {
        alert("There is no document to print.");
        return;
    }

    $('#modalPrint').on('shown.bs.modal', function () {
        var printLink = "/Home/Print?token=" + globalToken + "&printtotal=" + objctlDoc.TotalPages();
        $("#printFrame").attr("src", printLink);
    }).modal('show');

}

function ClosePrint() {
    $('#modalPrint').modal('hide');
}

function GoFS() {

    if (
        document.fullscreenElement ||
        document.webkitFullscreenElement ||
        document.mozFullScreenElement ||
        document.msFullscreenElement
    ) {
        if (document.exitFullscreen) {
            document.exitFullscreen();
        } else if (document.mozCancelFullScreen) {
            document.mozCancelFullScreen();
        } else if (document.webkitExitFullscreen) {
            document.webkitExitFullscreen();
        } else if (document.msExitFullscreen) {
            document.msExitFullscreen();
        }
    } else {
        var element = $('#divDocViewer').get(0);

        if (element.requestFullscreen) {
            element.requestFullscreen();
        } else if (element.mozRequestFullScreen) {
            element.mozRequestFullScreen();
        } else if (element.webkitRequestFullscreen) {
            element.webkitRequestFullscreen(Element.ALLOW_KEYBOARD_INPUT);
        } else if (element.msRequestFullscreen) {
            element.msRequestFullscreen();
        }
    }
}

function OpenUpload() {
    $('#myModal').modal('show');
}

$(document).ready(function () {

    // Open a default document (optional)
    OpenDocument('Sample.doc');

    $("#dropZoneForm").dropzone({
        url: "/Home/UploadFile",
        maxFiles: 1,
        paramName: "file",
        uploadMultiple: false,
        maxFilesize: 20,
        acceptedFiles:
            ".doc,.docx,.docm,.odt,.xls,.xlsx,.xlsm,.ods,.csv,.ppt,.pptx,.odp,.vsd,.vsdx,.mpp,.mppx,.pdf,.tif,.tiff,.dwg,.dxf,.dgn,.xps,.psd,.jpg,.jpeg,.jpe,.png,.bmp,.gif,.eml,.msg,.txt,.rtf,.xml,.epub,.svg,.html,.htm,.mht,.dcn,.dcm,.dng,.ico,.eps,.tga,.webp,.cdr,.cmx",
        addRemoveLinks: false,
        init: function () {
            var th = this;
            this.on("success",
                function (file, response) {
                    OpenDocument(response);
                    $('#myModal').modal('hide');
                }),
                this.on("error",
                    function (file, errorMessage, c) {
                        alert("Error uploading document [[" +
                            file.name +
                            "]]. Technical team has been notified.");
                    }),
                this.on("queuecomplete",
                    function () {
                        setTimeout(function () {
                            th.removeAllFiles();
                        },
                            3000);
                    });
        }
    });

});