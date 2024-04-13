// 
// Function to add change listener for file input
const addFileChangeListener = (inputId, spanId, stepNumber) => {
  const fileInput = document.getElementById(inputId);
  const spanElement = document.getElementById(spanId);

  fileInput.addEventListener('change', function() {
    if (fileInput.files.length > 0) {
      const fileName = fileInput.files[0].name;
      spanElement.textContent = fileName;
      markStepComplete(stepNumber); 
    } else {
      spanElement.textContent = '';
    }
  });
};

// Function to add change listener for date input
const addDateChangeListener = (inputId, stepNumber) => {
  const dateInput = document.getElementById(inputId);

  dateInput.addEventListener('change', function() {
    markDateInputComplete(inputId, stepNumber);
  });
};

// Function to mark step as complete
const markStepComplete = (stepNumber) => {
  const circle = document.querySelectorAll(`#step${stepNumber} .step-circle`);
  circle.forEach((circle) => {
    circle.innerHTML = `<img src="../static/image/correct.svg" alt="">`;
    circle.style.backgroundColor = "#81689D";
  });

  const stepLine = document.querySelector(`#step${stepNumber} .step-line`);
  stepLine.style.backgroundColor = "#81689D";
};

// Function to mark date input as complete
const markDateInputComplete = (inputId, stepNumber) => {
  const dateInput = document.getElementById(inputId);
  
  // Change background color to blue and text color to white
  dateInput.style.backgroundColor = "#81689D";
  dateInput.style.color = "white";

  // Mark the step as complete
  markStepComplete(stepNumber);
};

// Call addFileChangeListener for each file input field
addFileChangeListener('fileInput', 'fileInput', 1); 
addFileChangeListener('fileInput1', 'fileInput1', 2); 
addFileChangeListener('fileInput2', 'fileInput2', 3); 
addFileChangeListener('fileInput3', 'fileInput3', 4); 
addFileChangeListener('fileInput4', 'fileInput3', 5); 

addDateChangeListener('dateInput', 6); 

let currentStep = 1;
const steps = document.querySelectorAll(".step");

function showStep(stepNumber) {
  steps.forEach((step) => step.classList.remove("active"));
  document.getElementById(`step${stepNumber}`).classList.add("active");
}

// Open_Modal
$(document).ready(function () {
  $("#openModalBtn").click(function () {
    var proceedText = $(this).text(); // Fetching the text of the clicked element
    var condition = (proceedText === "proceed"); // Checking if the text is "proceed"
    
    if (condition) {
      // Clear the text if the condition is true
      $("#textElement").text("");
      // Show the modal if the condition is true
      $("#fileInfoModal").modal("show");
    } else {
      // Change the text to "proceed" if the condition is false
      $("#textElement").text("proceed");
    }
  });
});



$(document).ready(function () {
  // Initialize DataTable
  var table = $("#datatable").DataTable({
    responsive: true,
    dom:
      "<'row'<'col-sm-12 col-md-4 d-flex align-items-center'f><'col-sm-12 col-md-4 filter-buttons'><'col-sm-12 col-md-4 d-flex justify-content-end'B>>" +
      "<'row'<'col-sm-12'tr>>" +
      "<'row'<'col-sm-12 col-md-5'i><'col-sm-12 col-md-7'p>>",
    buttons: [
      {
        text: '<span>Delete All<img src="../static/image/delete.svg" alt="" class="ms-3"></span>',
        action: function (e, dt, node, config) {
          // Send a POST request to delete all files
          $.post("/delete_all/", function(data, status) {
            if (status === "success") {
              // Reload the page or update the table as needed
              location.reload(); // Example: Reload the page
            } else {
              // Handle errors
              console.error("Error deleting files:", data);
              alert("Error deleting files. Please try again.");
            }
          });
        },
        className: "btn-red",
        enabled: ($("#datatable tbody tr").length > 0), // Enable the button if there's data initially
      },
      {
        extend: "excelHtml5",
        text: '<span>Download<img src="../static/image/download.svg" alt="" class="ms-3"></span>',
        title: "Download",
        className: "",
      },
    ],
    language: {
      search: "",
      searchPlaceholder: "Search...",
    },
  });

  // Function to update delete button state
  function updateDeleteButtonState() {
    var hasData = table.rows().count() > 0;
    table.button(0).enable(hasData); // Enable/disable the button based on data presence
    if (!hasData) {
      // If there's no data, change button color to secondary
      $('.btn-red').addClass('btn btn-secondary').removeClass('btn-red');
    } else {
      // If there's data, change button color back to red
      $('.btn-secondary').addClass('btn-red').removeClass('btn btn-secondary');
    }
  }

  // Call the function initially
  updateDeleteButtonState();

  // Listen to draw event to update button state
  table.on('draw', function () {
    updateDeleteButtonState();
  });
});
