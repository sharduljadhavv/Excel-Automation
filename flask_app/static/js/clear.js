// Variable to track if the button text has been cleared
let isTextCleared = false;

// Function to handle "Proceed" button click
function handleProceed() {
    
    let button = document.getElementById("openModalBtn");

    if (!isTextCleared) {
        button.innerText = "clear All Files";
        button.classList.add("btn-red");
        $('#fileInfoModal').modal('show');
        isTextCleared = true;
    } else {
        button.innerText = "Proceed";
        button.classList.remove("btn-red"); 
        isTextCleared = false;
    }
}

// Function to handle file download (assuming it opens modal on click)
function handleFile() {
    
    $('#fileInfoModal').modal('show');
}
