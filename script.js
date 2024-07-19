document.getElementById('uploadForm').addEventListener('submit', function(event) {
    event.preventDefault(); 
    let fileInput = document.getElementById('fileUpload');
    let ageThreshold = parseInt(document.getElementById('ageThreshold').value, 10);
    let file = fileInput.files[0];
    if (file) {
        let reader = new FileReader();
        reader.onload = function(e) {
            let contents = e.target.result;
            if (file.name.endsWith('.csv')) {
                parseCSV(contents, ageThreshold);
            } else if (file.name.endsWith('.xlsx')) {
                try {
                    parseXLSX(contents, ageThreshold);
                } catch (error) {
                    alert('Error parsing XLSX file. Please ensure the file is a valid XLSX file.');
                }
            } else if (file.name.endsWith('.docx')) {
                try {
                    parseDOCX(contents, ageThreshold);
                } catch (error) {
                    alert('Error parsing DOCX file. Please ensure the file is a valid DOCX file.');
                }
            } else {
                alert('Unsupported file format. Please upload a CSV, XLSX, or DOCX file.');
            }
        };
        if (file.name.endsWith('.csv')) {
            reader.readAsText(file);
        } else if (file.name.endsWith('.xlsx')) {
            reader.readAsText(file);
        } else if (file.name.endsWith('.docx')) {
            reader.readAsText(file);
        }
    }
});
function parseCSV(contents, ageThreshold) {
    let lines = contents.split('\n');
    let results = [];   
    for (let line of lines) {
        let data = line.split(',');
        if (data.length >= 2) { 
            let name = data[0].trim();
            let age = parseInt(data[1].trim(), 10);
            
            if (!isNaN(age) && age >= ageThreshold) {
                results.push({ name: name, age: age });
            }
        }
    }   
    displayResults(results);
    enableDownload(results);
}
function parseXLSX(contents, ageThreshold) {
    let workbook = XLSX.read(contents, { type: 'array' });
    let firstSheetName = workbook.SheetNames[0];
    let worksheet = workbook.Sheets[firstSheetName];
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    let results = [];
    for (let i = 1; i < jsonData.length; i++) {
        let row = jsonData[i];
        if (row.length >= 2) {
            let name = row[0].trim();
            let age = parseInt(row[1].trim(), 10);
            
            if (!isNaN(age) && age >= ageThreshold) {
                results.push({ name: name, age: age });
            }
        }
    }
    displayResults(results);
    enableDownload(results);
}
function parseDOCX(contents, ageThreshold) {
    let zip = new PizZip(contents);
    let doc = new window.docxtemplater(zip);
    let text = doc.getFullText();
    let lines = text.split('\n');
    let results = [];   
    for (let line of lines) {
        let data = line.split('\t');
        if (data.length >= 2) {
            let name = data[0].trim();
            let age = parseInt(data[1].trim(), 10);           
            if (!isNaN(age) && age >= ageThreshold) {
                results.push({ name: name, age: age });
            }
        }
    } 
    displayResults(results);
    enableDownload(results);
}
function displayResults(results) {
    let resultArea = document.getElementById('resultArea');
    resultArea.innerHTML = '';   
    if (results.length > 0) {
        let list = document.createElement('ul');     
        results.forEach(function(person) {
            let listItem = document.createElement('li');
            listItem.textContent = ` Name: ${person.name} - Age: ${person.age}`;
            list.appendChild(listItem);
        });     
        resultArea.appendChild(list);
    } else {
        resultArea.textContent = 'No senior citizens found.';
    }
}
function enableDownload(results) {
    let downloadButton = document.getElementById('downloadButton');
    downloadButton.style.display = 'block';
    downloadButton.onclick = function() {
        let wb = XLSX.utils.book_new();
        let ws = XLSX.utils.json_to_sheet(results);
        XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');
        XLSX.writeFile(wb, 'filtered_senior_citizens.xlsx');
    };
}
