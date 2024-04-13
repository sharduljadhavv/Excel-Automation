                $(document).ready(function () {
                    $('#viewDetailsBtn').click(function () {
                        $('#popupTable').fadeIn();
                        $('#backdrop').fadeIn();
                    });
                
                    $('.close ').click(function () {
                        $('#popupTable').fadeOut();
                        $('#backdrop').fadeOut();
                    });
                    $('.buttons ').click(function () {
                        $('#popupTable').fadeOut();
                        $('#backdrop').fadeOut();
                    });
                });
                
                document.getElementById("viewDetailsBtn").addEventListener("click", function() {
            document.getElementById("popupTable").style.display = "block";
            });
            
            function closePopup() {
            document.getElementById("popupTable").style.display = "none";
            } 



        var fromDate;
        $('#fromDate').on('change',function(){
          fromDate = $(this).val();
          $('#toDate').prop('min',function(){
            return fromDate;
    
          })
        });
    
        var toDate;
        $('#toDate').on('change',function(){
          toDate = $(this).val();
          $('#fromDate').prop('max',function(){
            return toDate;
    
          })
        });