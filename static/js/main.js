function checkForm(event) {
    event.preventDefault();

    const surnameField = $("#surname_field");
     const name_field = $("#name_field");


    if (surnameField.val() === "") {
        surnameField.addClass("invalid");
    } else {
        surnameField.removeClass("invalid");
    }
}

const form = document.querySelector("#form");
form.addEventListener('submit', checkForm);
