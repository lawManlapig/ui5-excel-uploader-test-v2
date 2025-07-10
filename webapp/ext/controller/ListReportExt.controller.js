sap.ui.define([
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/core/Fragment",
    "xlsx"
], function (MessageToast, MessageBox, Fragment, XLSX) {
    'use strict';

    // Controller
    function _createUploadController(extensionAPI) {
        let uploadDialog;

        // Internal Functions
        function closeDialog() {
            uploadDialog && uploadDialog.close()
        }

        function setOkButtonEnabled(isEnabled) {
            uploadDialog && uploadDialog.getButtons()[0].setEnabled(isEnabled);
        }

        function showError(message) {
            MessageBox.error("Upload failed: " + message, { title: "Error" });
        }

        return {
            // Dialog Handlers
            onBeforeOpen: function (event) {
                uploadDialog = event.getSource();
                // extensionAPI.addDependent(uploadDialog);
            },

            onAfterClose: function (event) {
                // extensionAPI.removeDependent(uploadDialog);
                uploadDialog.destroy();
                uploadDialog = undefined;
            },

            onChange: function (event) {
                //Information Message
                MessageBox.information("Make sure you are using the correct template before uploading.");
            },

            // Button Handlers
            onCancelPress: function (event) {
                closeDialog();
            },

            onUploadPress: function (event) {
                MessageToast.show("Upload Button click invoked.");
            },

            onTemplateDownloadPress: function (event) {
                MessageToast.show("Download Template Button click invoked.");
            },

            // File Uploader Handlers
            onUploadComplete: function (event) {
                MessageToast.show("Download Template Button click invoked.");
            },

            onFileEmpty: function (event) {
                setOkButtonEnabled(false)
            },

            onFileAllowed: function (event) {
                setOkButtonEnabled(true)
            },

            onTypeMismatch: function (event) {
                let supportedFileTypes = event
                    .getSource()
                    .getFileType()
                    .map(function (fileType) {
                        return "*." + fileType;
                    })
                    .join(", ");

                showError(
                    "The file type *." +
                    event.getParameter("fileType") +
                    " is not supported. Choose one of the following types: " +
                    supportedFileTypes
                );
            },
        }
    }
    // View Handlers
    return {
        openExcelUploadDialog: function (oEvent) {
            console.log(XLSX.version);
            let view = this.getView();

            Fragment.load({
                id: "excel_upload",
                name: "lawbuildingssample.ext.fragment.ExcelUpload",
                type: "XML",
                controller: _createUploadController(this)
            }).then((inDialog) => {
                let fileUploader = Fragment.byId("excel_upload", "fileUploader");

                this.dialog = inDialog;
                this.dialog.open();
            }).catch(error => alert(error.message));
        },
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