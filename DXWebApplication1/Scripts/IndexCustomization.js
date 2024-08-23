    function onDocumentLoaded(s, e) {
        console.log("Hello World");
        s.UpdateAllFields();
    }

    function onInit(s, e) {
        console.log("RichEdit Initialized");
        s.commands.updateAllFields.execute();
    }
    function OnProtectDocumentClick(s, e) {
        RichEdit.PerformCallback({ actioName: "protectDocumentFields" });
    }

    function OnUpdateProtectedFields(s, e) {
        RichEdit.PerformCallback({ actioName: "updateProtectedFields" });
    }
    window.onInit = onInit;
    window.onDocumentLoaded = onDocumentLoaded;
