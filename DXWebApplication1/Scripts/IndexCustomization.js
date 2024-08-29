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

    function OnProtectSection(s, e) {
        RichEdit.PerformCallback({ actioName: "protectSection" });
    }

    window.onInit = onInit;
    window.onDocumentLoaded = onDocumentLoaded;
