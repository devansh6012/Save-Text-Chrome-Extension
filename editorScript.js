document.querySelector('#saveBtn').addEventListener('click', function () {
    console.log('hello');
    try {
        const editedText = document.getElementById('editor').innerText;
        saveAsDocFile(editedText);
        window.close();
    } catch (error) {
        console.error(error);
    }
});

function saveAsDocFile(text) {
    try {
        const docContent = '<html xmlns:w="urn:schemas-microsoft-com:office:word">' +
            '<head>' +
            '<style>' +
            '<!-- Add your desired styles here -->' +
            '</style>' +
            '</head>' +
            '<body>' +
            '<div>' + text + '</div>' +
            '</body>' +
            '</html>';

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
    } catch (error) {
        console.error(error);
    }
}
