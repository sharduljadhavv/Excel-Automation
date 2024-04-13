function handleFile() {
  const fileInput = document.getElementById('fileInput');
  

  // Check if a file is selected
  if (!fileInput.files.length) {
    alert('Please select a file');
    return;
  }

  const file = fileInput.files[0];

  // Check if the file is of Excel type
  if (!file.name.endsWith('.xls') && !file.name.endsWith('.xlsx')) {
    alert('Please select an Excel file');
    return;
  }

  const reader = new FileReader();
  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Process the workbook here, for simplicity let's just download it again
    const processedData = XLSX.write(workbook, { type: 'array' });
    const blob = new Blob([processedData], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'processed_file.xlsx';
    a.style.display = 'none'; // Hide the anchor element
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  reader.readAsArrayBuffer(file);
}
