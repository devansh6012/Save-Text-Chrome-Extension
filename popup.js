const btn = document.querySelector('#saveButton');

document.querySelector('#editButton').addEventListener('click', async () => {
    let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    console.log('edit btn');
    chrome.scripting.executeScript({
        target: { tabId: tab.id },
        function: openRichTextEditor,
    });
});


btn.addEventListener('click', async () => {
    console.log('clicked');
    let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    console.log(tab.id);
    chrome.scripting.executeScript({
        target: { tabId: tab.id },
        function: pickText,
    })
})

function pickText(){
    try {
        const selObj = window.getSelection();
        // window.alert(selObj);

        let text = selObj.toString();
        console.log(text);
        
        
        const docContent = `
        <html xmlns:w="urn:schemas-microsoft-com:office:word">
            <head>
                <style>
                    <!-- Add your desired styles here -->
                </style>
            </head>
            <body>
                <div>${text}</div>
            </body>
        </html>`;

    const blob = new Blob(['\ufeff', docContent], {
        type: 'application/msword',
    });

    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'selectedText.doc';
    document.body.appendChild(a);

    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    } catch (error) {
        console.error(error);
    }
}


function openRichTextEditor() {
    const selectedText = window.getSelection().toString();

    const editorWindow = window.open('', 'Rich Text Editor', 'width=600,height=400');
    console.log(chrome.runtime.getURL('editorScript.js'));

    editorWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Rich Text Editor</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                }
                div[contenteditable] {
                    border: 1px solid #ccc;
                    padding: 10px;
                    min-height: 200px;
                }
            </style>
        </head>
        <body>
            <div contenteditable="true" id="editor">${selectedText}</div>
            <br>
            <button id="saveBtn">Save as .doc</button>
            <script src="${chrome.runtime.getURL('editorScript.js')}"></script>
        </body>
        </html>
    `);
}


function saveAsDocFile(text) {
    const docContent = `
        <html xmlns:w="urn:schemas-microsoft-com:office:word">
            <head>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                    }
                </style>
            </head>
            <body>
                <div>${text}</div>
            </body>
        </html>`;

    const blob = new Blob(['\ufeff', docContent], {
        type: 'application/msword',
    });

    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'editedText.doc';
    document.body.appendChild(a);

    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
