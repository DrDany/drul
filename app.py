from flask import Flask, redirect, url_for, render_template, request
import exel_handler

# Flask
app = Flask(__name__)


@app.route("/", methods=('GET', 'POST'))
def comment():
    '''
    Create new comment
    '''

    if request.method == 'GET':
        return render_template('web/new_comment.html')
    else:
        surname = request.form["surname"]
        name = request.form["name"]
        birthdate = request.form["birthdate"]
        patranomic = request.form["patranomic"]
        citizen = request.form["citizen"]
        gender = request.form["gender"]
        # birth_place = request.form["birth_place"]
        # birth_city = request.form["birth_city"]
        # doc_type = request.form["doc_type"]
        doc_seria = request.form["doc_seria"]
        doc_number = request.form["doc_number"]
        doc_date = request.form["doc_date"]
        doc_end = request.form["doc_end"]
        profession = request.form["profession"]
        date_income = request.form["date_income"]
        date_stay_to = request.form["date_stay_to"]
        # region = request.form["region"]
        # city = request.form["city"]
        # district = request.form["district"]
        # street = request.form["street"]
        # street_number = request.form["street_number"]
        # flat_number = request.form["flat_number"]

        mig_card_ser = request.form["mig_card_ser"]
        mig_card_number = request.form["mig_card_number"]
        # mig_card_region = request.form["mig_card_region"]
        # mig_card_city = request.form["mig_card_city"]
        # mig_card_street_number = request.form["mig_card_street_number"]
        surname_host = request.form["surname_host"]
        name_host = request.form["name_host"]
        patr_host = request.form["patr_host"]
        # date_host_birth = request.form["date_host_birth"]
        host_doc_seria = request.form["host_doc_seria"]
        host_doc_number = request.form["host_doc_number"]
        date_host_pass = request.form["date_host_pass"]

        str1 = request.form["str1"]
        str2 = request.form["str2"]
        str3 = request.form["str3"]
        str4 = request.form["str4"]
        exel_handler.add_new_exel(surname, name, patranomic, citizen, birthdate, gender, doc_seria, doc_number, doc_date, doc_end, profession, date_income, date_stay_to, mig_card_ser, mig_card_number, surname_host, name_host, patr_host, host_doc_seria, host_doc_number, date_host_pass, str1, str2, str3, str4)
        return redirect(url_for('comment'))


if __name__ == "__main__":
    app.run()
