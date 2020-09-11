$('[data-toggle=confirmation]').confirmation({
    rootSelector: '[data-toggle=confirmation]'
    //btnOkLabel: 'Si',
    //btnCancelLabel: 'No'
});

function setDate() {
        $('.datepicker').datepicker()
  }

function checkForm() {

    var surnameField = $("#surname_field");
    var nameField = $("#name_field");
    var commentFiled = $("#comment_field");
    var emailField = $("#email_filed");
    var phoneFiled = $("#phone_field");

    var emailPattern = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/;
    var phonePattern = /^[(]\d{1,6}[)]\d{7,9}/;
    var isEmail = emailPattern.test(emailField.val());
    var isCorrectPhone = phonePattern.test(phoneFiled.val());

    if (isEmail) {
        emailField.removeClass("invalid");
    } else {
        emailField.addClass("invalid");
    }

    if (isCorrectPhone) {
        phoneFiled.removeClass("invalid");
    } else {
        phoneFiled.addClass("invalid");
    }

    if (surnameField.val() === "") {
        surnameField.addClass("invalid");
    } else {
        surnameField.removeClass("invalid");
    }
    if (nameField.val() === "") {
        nameField.addClass("invalid");
    } else {
        nameField.removeClass("invalid");
    }

    if (commentFiled.val() === "") {
        commentFiled.addClass("invalid");
        return false;
    } else {
        commentFiled.removeClass("invalid");
    }
    return true;
}