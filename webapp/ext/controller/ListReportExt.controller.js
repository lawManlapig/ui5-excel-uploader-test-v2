sap.ui.define([
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/core/Fragment"
], function (MessageToast, MessageBox, Fragment) {
    'use strict';

    function _createUploadController(oExtensionAPI) {
        //Local Variable
        let oUploadDialog;

        function setOkButtonEnabled(bIsEnabled) {
            oUploadDialog && oUploadDialog.getBeginButton().setEnabled(bIsEnabled);
        }

        function setDialogBusy(bIsBusy) {
            oUploadDialog.setBusy(bIsBusy)
        }

        function closeDialog() {
            oUploadDialog && oUploadDialog.close()
        }

        function showError(sMessage) {
            MessageBox.error("Upload failed: " + sMessage, { title: "Error" });
        }

        function byId(sId) {
            return sap.ui.core.Fragment.byId("excelUploadDialog", sId);
        }

        //File Uploader Handlers
        return {
            //Dialog Handler
            onBeforeOpen: function (oEvent) {
                oUploadDialog = oEvent.getSource();
                oExtensionAPI.addDependent(oUploadDialog);
            },

            //Dialog Handler
            onAfterClose: function (oEvent) {
                oExtensionAPI.removeDependent(oUploadDialog);
                oUploadDialog.destroy();
                oUploadDialog = undefined;
            },

            //Dialog Handler
            onChange: function (oEvent) {
                //Information Message
                MessageBox.information("Make sure you are using the correct template before uploading.");
            },

            //Button Handler
            onUpload: function (oEvent) {
                //Set the dialog to 'busy'
                setDialogBusy(true);

                //Set Header Parameters
                var oFileUploader = byId("fileUploader"),
                    headPar = new sap.ui.unified.FileUploaderParameter();

                headPar.setName('slug');
                headPar.setValue(sEntity);
                oFileUploader.removeHeaderParameter('slug');
                oFileUploader.addHeaderParameter(headPar);

                //Set URI
                var sUploadUri = oExtensionAPI._controller.extensionAPI._controller._oAppComponent
                    .getManifestObject().resolveUri('../../odata/v4/upload-utility-meters-srv/ExcelUpload/excel');

                oFileUploader.setUploadUrl(sUploadUri);

                //Validate Upload File
                oFileUploader
                    .checkFileReadable().then(function () {
                        oFileUploader.upload();
                    }).catch(function (error) {
                        showError("The file cannot be read.");
                        setDialogBusy(false)
                    });
            },

            //Button Handler
            onCancel: function (oEvent) {
                closeDialog();
            },

            //Dialog Handler
            onTypeMismatch: function (oEvent) {
                var sSupportedFileTypes = oEvent
                    .getSource()
                    .getFileType()
                    .map(function (sFileType) {
                        return "*." + sFileType;
                    })
                    .join(", ");

                showError(
                    "The file type *." +
                    oEvent.getParameter("fileType") +
                    " is not supported. Choose one of the following types: " +
                    sSupportedFileTypes
                );
            },

            //Dialog Handler
            onFileAllowed: function (oEvent) {
                setOkButtonEnabled(true)
            },

            //Dialog Handler
            onFileEmpty: function (oEvent) {
                setOkButtonEnabled(false)
            },

            //Dialog Handler
            onUploadComplete: function (oEvent) {
                var sMessageRaw = oEvent.getParameter("responseRaw"),
                    iStatus = oEvent.getParameter("status"),
                    oFileUploader = oEvent.getSource();

                oFileUploader.clear();
                setOkButtonEnabled(false)
                setDialogBusy(false)

                if (iStatus >= 400) {
                    var oResponse;

                    try {
                        oRawResponse = JSON.parse(sMessageRaw);
                        showError(oResponse);
                    } catch (e) {      //For XML Responses
                        if (window.DOMParser) {
                            var oParser = new DOMParser(),
                                sXMLDoc = oParser.parseFromString(sMessageRaw, "text/html");
                        } else {
                            sXMLDoc = new ActiveXObject("Microsoft.XMLDOM");
                            sXMLDoc.async = false;
                            sXMLDoc.loadXML(sMessageRaw);
                        };

                        try {
                            showError(sXMLDoc.getElementsByTagName("pre")[0].childNodes[0].nodeValue);
                        } catch (e) {
                            showError('Something went wrong.');
                        };
                    };
                } else {
                    MessageToast.show("File uploaded successfully");
                    oExtensionAPI.refresh();
                    closeDialog();
                }
            }
        }
    }

    // View Handlers
    return {
        openExcelUploadDialog: function (oEvent) {
            // MessageToast.show("Custom handler invoked.");
            let oSelf = this;

            //Opens the custom fragment created...
            this.loadFragment({
                id: 'excelUploadDialog',
                name: 'lawbuildingssample.ext.fragment.ExcelUpload', //Filepath of the fragment
                controller: _createUploadController(oSelf)
            }).then(function (oDialog) {
                oDialog.open();
            });
        }
    };
});

/* global XLSX:true 
sap.ui.define(["sap/m/MessageToast", "sap/m/MessageBox"],
    function (MessageToast, MessageBox) {
        "use strict";
        function _createUploadController(sEntity, oExtensionAPI,) {
            //Local Variable
            var oUploadDialog;
 
            function setOkButtonEnabled(bIsEnabled) {
                oUploadDialog && oUploadDialog.getBeginButton().setEnabled(bIsEnabled);
            }
 
            function setDialogBusy(bIsBusy) {
                oUploadDialog.setBusy(bIsBusy)
            }
 
            function closeDialog() {
                oUploadDialog && oUploadDialog.close()
            }
 
            function showError(sMessage) {
                MessageBox.error("Upload failed: " + sMessage, { title: "Error" });
            }
 
            function byId(sId) {
                return sap.ui.core.Fragment.byId("excelUploadDialog", sId);
            }
 
            //File Uploader Handlers
            return {
                //Dialog Handler
                onBeforeOpen: function (oEvent) {
                    oUploadDialog = oEvent.getSource();
                    oExtensionAPI.addDependent(oUploadDialog);
                },
 
                //Dialog Handler
                onAfterClose: function (oEvent) {
                    oExtensionAPI.removeDependent(oUploadDialog);
                    oUploadDialog.destroy();
                    oUploadDialog = undefined;
                },
 
                //Dialog Handler
                onChange: function (oEvent) {
                    //Information Message
                    MessageBox.information("Make sure you are using the correct template before uploading.");
                },
 
                //Button Handler
                onUpload: function (oEvent) {
                    //Set the dialog to 'busy'
                    setDialogBusy(true);
 
                    //Set Header Parameters
                    var oFileUploader = byId("fileUploader"),
                        headPar = new sap.ui.unified.FileUploaderParameter();
 
                    headPar.setName('slug');
                    headPar.setValue(sEntity);
                    oFileUploader.removeHeaderParameter('slug');
                    oFileUploader.addHeaderParameter(headPar);
 
                    //Set URI
                    var sUploadUri = oExtensionAPI._controller.extensionAPI._controller._oAppComponent
                        .getManifestObject().resolveUri('../../odata/v4/upload-utility-meters-srv/ExcelUpload/excel');
 
                    oFileUploader.setUploadUrl(sUploadUri);
 
                    //Validate Upload File
                    oFileUploader
                        .checkFileReadable().then(function () {
                            oFileUploader.upload();
                        }).catch(function (error) {
                            showError("The file cannot be read.");
                            setDialogBusy(false)
                        });
                },
 
                //Button Handler
                onCancel: function (oEvent) {
                    closeDialog();
                },
 
                //Dialog Handler
                onTypeMismatch: function (oEvent) {
                    var sSupportedFileTypes = oEvent
                        .getSource()
                        .getFileType()
                        .map(function (sFileType) {
                            return "*." + sFileType;
                        })
                        .join(", ");
 
                    showError(
                        "The file type *." +
                        oEvent.getParameter("fileType") +
                        " is not supported. Choose one of the following types: " +
                        sSupportedFileTypes
                    );
                },
 
                //Dialog Handler
                onFileAllowed: function (oEvent) {
                    setOkButtonEnabled(true)
                },
 
                //Dialog Handler
                onFileEmpty: function (oEvent) {
                    setOkButtonEnabled(false)
                },
 
                //Dialog Handler
                onUploadComplete: function (oEvent) {
                    var sMessageRaw = oEvent.getParameter("responseRaw"),
                        iStatus = oEvent.getParameter("status"),
                        oFileUploader = oEvent.getSource();
 
                    oFileUploader.clear();
                    setOkButtonEnabled(false)
                    setDialogBusy(false)
 
                    if (iStatus >= 400) {
                        var oResponse;
 
                        try {
                            oRawResponse = JSON.parse(sMessageRaw);
                            showError(oResponse);
                        } catch (e) {      //For XML Responses
                            if (window.DOMParser) {
                                var oParser = new DOMParser(),
                                    sXMLDoc = oParser.parseFromString(sMessageRaw, "text/html");
                            } else {
                                sXMLDoc = new ActiveXObject("Microsoft.XMLDOM");
                                sXMLDoc.async = false;
                                sXMLDoc.loadXML(sMessageRaw);
                            };
 
                            try {
                                showError(sXMLDoc.getElementsByTagName("pre")[0].childNodes[0].nodeValue);
                            } catch (e) {
                                showError('Something went wrong.');
                            };
                        };
                    } else {
                        MessageToast.show("File uploaded successfully");
                        oExtensionAPI.refresh();
                        closeDialog();
                    }
                }
            }
        }
        //View Handlers
        return {
            //Upload File
            onUploadFile: function (oEvent) {
                let oSelf = this;
 
                //Opens the custom fragment created...
                this.loadFragment({
                    id: 'excelUploadDialog',
                    name: 'umupload.extensions.fragments.uploadxls', //Filepath of the fragment
                    controller: _createUploadController('umm_db_UtilityMeterUpload', oSelf)
                }).then(function (oDialog) {
                    oDialog.open();
                });
            },
 
            //Download Template
            onDownloadTemp: function (oEvent) {
                var excelColumns = [];
 
                //Set the Headers/Guide in the Excel File as JSON
                var jsonHeaders = { 
                    COMPANYCODE_COMPANYCODE: 'Company Code',
                    PROJECTCODE_PROJECTCODE: 'Project Code',
                    BUILDING_BUILDINGCODE: 'Building Code',
                    LAND_LANDCODE: 'Land Code',
                    // METERNUMBER: '<Just leave this as blank>',
                    TYPEOFMETER_CODE: 'Type Of Meter',
                    METEREQBRAND: 'Meter Equipment Brand',
                    SERIALNUMBER: 'Serial Number',
                    MULTIPLIER: 'Multiplier',
                    REALESTATEOBJECT_REOBJECT: 'Real Estate Object',
                    COUNTEROVRFLW: 'Counter Overflow',
                    UOM_ITEMCODE: 'Unit of Measure',
                    // CREATEDAT: '<Just leave this as blank>',
                    // CREATEDBY: '<Just leave this as blank>',
                    // MODIFIEDAT: '<Just leave this as blank>',
                    // MODIFIEDBY: '<Just leave this as blank>',
                    // ID: '<Just leave this as blank>',
                    // STATUS: '<Just leave this as blank>',
                    // MESSAGEINFO: '<Just leave this as blank>'
                };
 
                //Add to Array
                excelColumns.push(jsonHeaders);
 
                //Worksheeet and Workbook
                const ws = XLSX.utils.json_to_sheet(excelColumns);
                
                //Sets column width
                ws['!cols'] = [
                    {wch:30}, {wch:30}, {wch:30}, {wch:30}, {wch:30},
                    {wch:30}, {wch:30}, {wch:30}, {wch:30}, {wch:30},
                    {wch:30}
                ];
 
                const wb = XLSX.utils.book_new();
 
                //Add to Sheet
                XLSX.utils.book_append_sheet(wb, ws, 'Utility Meters Upload');
 
                //Download File
                XLSX.writeFile(wb, 'Utility Meters Upload Template.xlsx');
 
                //Flavor Text
                MessageToast.show("Template File Downloading...");
            }
        };
    });

*/