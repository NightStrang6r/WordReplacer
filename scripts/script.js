const zip = new JSZip();

function main() {
    const inputFile = document.querySelector('.js-file-input');
    inputFile.addEventListener('input', async (e) => onFileInput(e));
}

async function onFileInput(event) {
    try {
        const file = event.target.files[0];
        
        if(!file.name || !(file.name.endsWith('.docx') || file.name.endsWith('.doc'))) {
            displayMessage('File is not a .docx or .doc file');
            return;
        }

        const content = await zip.loadAsync(file);
        const files = Object.keys(content.files);

        if(!files || !files.length) {
            displayMessage('This document file is not supported');
            return;
        }

        let documentXML = null;
        for (const file of files) {
            if(file != "word/document.xml") continue;
            documentXML = await zip.file(file).async('string');
        }

        if(!documentXML) {
            displayMessage('This document file is not supported');
            return;
        }

        const vars = getVariables(documentXML);

        if(!vars || !vars.length) {
            displayMessage('This document file does not contain any variables');
            return;
        }

        for (const variable of vars) {
            addNewTR(variable);
        }
    } catch (e) {
        displayMessage(`This document file is not supported: ${e}`);
    }
}

function getVariables(stringXML) {
    const variables = [];
    const regex = /{.*?}/g;
    let m;

    while ((m = regex.exec(stringXML)) !== null) {
        if (m.index === regex.lastIndex) {
            regex.lastIndex++;
        }

        m.forEach((match) => {
            variables.push(match);
        });
    }

    console.log(variables);
    return variables;
}

function displayMessage(message) {
    const messageContainer = document.querySelector('.js-result');
    messageContainer.innerHTML = message;

    setTimeout(() => {
        messageContainer.innerHTML = '';
    }, 3000);
}

function addNewTR(name) {
    const table = document.querySelector('.js-table');
    const row = table.insertRow(-1);
    const cell1 = row.insertCell(0);
    const cell2 = row.insertCell(1);
    cell1.innerHTML = name;
    cell2.innerHTML = `<input type="text" class="js-input" name="${name}" />`;
}

/*

const content = await zip.loadAsync(file);

const files = Object.keys(content.files);

for (const file of files) {
    if(file != "word/document.xml") continue;
    const content = await zip.file(file).async('string');
    const newContent = content.replace("UNIT 13", "NOT UNIT");
    zip.file(file, newContent);

    const blob = await zip.generateAsync({type: "blob"});
    saveAs(blob, "newFile.docx");
}

*/

main();