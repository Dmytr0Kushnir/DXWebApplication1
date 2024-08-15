    function onDocumentLoaded(s, e) {
        console.log("Hello World");
        s.UpdateAllFields();
    }

    function onInit(s, e) {
        console.log("RichEdit Initialized");
        s.commands.updateAllFields.execute();
    }

    window.onInit = onInit;
    window.onDocumentLoaded = onDocumentLoaded;
