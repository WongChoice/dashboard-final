var checkbox = document.querySelector("input[type=checkbox][name=January][id=firstCheckbox]");

checkbox.addEventListener('change', function() {
  if (this.checked) {
     worksheetname2 = 'March';
    location.reload();
    console.log("Checkbox is not checked..");
  } else {
    console.log("Checkbox is not checked..");
  }
});