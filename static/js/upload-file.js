$('#reference-file').on('click', function(event) {
    clickOnInput(event);
})
$('#original-file').on('click', function(event) {
    clickOnInput(event);
})
$('#reference-file').on('change', function(event) {
    getFileName(event)
})
$('#original-file').on('change', function(event) {
    getFileName(event)
})
function clickOnInput(event) {
    const targetId = event.currentTarget.id;
    document.querySelector('#' + targetId + ' input').click();
}

function getFileName(event) {
    var name = event.target.files[0].name;
    const targetId = event.currentTarget.id;
    console.log(name);
    $('#' + targetId + ' p').text(name);
}

/*$(document).ready(function() {
         $('#upload-form').submit(function(event) {
            event.preventDefault();
            $.ajax({
               type: 'POST',
               url: '/upload',
               success: function() {
                  alert('Form submitted!');
               }
            });
         });
      });*/
  