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
        let internalExtensionAPI;
        let excelDataArray = [];

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
                internalExtensionAPI = extensionAPI;
            },

            onAfterClose: function (event) {
                // extensionAPI.removeDependent(uploadDialog);
                uploadDialog.destroy();
                uploadDialog = undefined;
            },

            onChange: function (event) {
                //Information Message
                MessageBox.information("Make sure you are using the correct template before uploading.");

                let fileUploader = Fragment.byId("excel_upload", "fileUploader"); // Get File Uploader instance
                let file = jQuery.sap.domById(fileUploader.getId() + "-fu").files[0]; // Get File details
                let reader = new FileReader();

                // Push to global array
                reader.onload = (e) => {
                    let xlsx_content = e.currentTarget.result;
                    let workbook = XLSX.read(xlsx_content, { type: 'binary' });
                    let excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets["Sheet1"]);

                    workbook.SheetNames.forEach(function (sheetName) {
                        // Appending the excel file data to the global variable
                        // that.excelSheetsData.push(XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]));
                        excelDataArray.push(XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]));
                    });
                };

                reader.readAsArrayBuffer(file);
            },

            // Button Handlers
            onCancelPress: function (event) {
                closeDialog();
            },

            onUploadPress: function (event) {
                let that = this;

                // Creating a promise as the extension api accepts odata call in form of promise only
                let addMessage = function () {
                    return new Promise((resolve, reject) => {
                        that.callOdata(resolve, reject);
                    });
                };

                // Build arguments then call the OData Service
                let parameters = { sActionLabel: event.getSource().getText() };
                internalExtensionAPI.extensionAPI.securedExecution(addMessage, parameters);

                closeDialog();
            },

            onTemplateDownloadPress: function (event) {
                let excelColumnList = [];
                let columnList = {};

                let model = internalExtensionAPI.getView().getModel();
                let building = model.getServiceMetadata().dataServices.schema[0].entityType.find(x => x.name === 'BuildingsType');

                // Set the list of entity property, that has to be present in excel file template
                let propertyList = ['BuildingId', 'BuildingName', 'NRooms', 'AddressLine',
                    'City', 'State', 'Country'];

                // Finding the property description corresponding to the property id
                propertyList.forEach((value, index) => {
                    let label = "";
                    let property = building.property.find(x => x.name === value);

                    // Logic for label :
                    // Get the label from extensions node... but if not found, separate the "value" variable by spaces
                    label = property.extensions?.find(x => x.name === 'label')?.value || value.replace(/([A-Z])/g, ' $1').trim();
                    columnList[label] = '';
                });
                excelColumnList.push(columnList);

                // Initialising the excel work sheet
                const ws = XLSX.utils.json_to_sheet(excelColumnList);
                // Creating the new excel work book
                const wb = XLSX.utils.book_new();
                // Set the file value
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                // Download the created excel file
                XLSX.writeFile(wb, 'RAP - Buildings.xlsx');

                MessageToast.show("Template File Downloading...");
            },

            // File Uploader Handlers
            onUploadComplete: function (event) {
                // getting the UploadSet Control reference
                let fileUploader = Fragment.byId("excel_upload", "fileUploader");
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

            // Helper
            callOdata: function (resolve, reject) {
                let model = internalExtensionAPI.getView().getModel();
                let payload = {};

                excelDataArray[0].forEach((value, index) => {
                    // Setting Payload Data
                    payload = {
                        "BuildingName": value["Building Name"],
                        "NRooms": value["Number of rooms"],
                        "AddressLine": value["Address Line"],
                        "City": value["City"],
                        "State": value["State"],
                        "Country": value["Country"]
                    };

                    // Setting excel file row number for identifying the exact row in case of error or success
                    payload.ExcelRowNumber = (index + 1);

                    // Calling the odata service
                    model.create("/Buildings", payload, {
                        success: (result) => {
                            console.log(result);
                            let messageManager = sap.ui.getCore().getMessageManager();
                            let message = new sap.ui.core.message.Message({
                                message: "Building Created with ID: " + result.BuildingId,
                                persistent: true, // Create message as transition message
                                type: sap.ui.core.MessageType.Success
                            });
                            messageManager.addMessages(message);
                            resolve();
                        },
                        error: reject
                    });
                });
            }
        }
    }
    // View Handlers
    return {
        openExcelUploadDialog: function (oEvent) {
            console.log(XLSX.version);
            let view = this.getView();

            Fragment.load({
                id: "excel_upload",
                name: "demobuildings.ext.fragment.ExcelUpload",
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