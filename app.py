from flask import Flask, redirect, url_for, render_template, request, flash
import models as db_handler
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
        # passport = request.form["passport"]
        citizen = request.form["citizen"]
        birth_place = request.form["birth_place"]
        birth_city = request.form["birth_city"]
        doc_type = request.form["doc_type"]
        doc_seria = request.form["doc_seria"]
        doc_number = request.form["doc_number"]
        doc_date = request.form["doc_date"]
        doc_end = request.form["doc_end"]
        exel_handler.add_new_exel(surname, name, birthdate, citizen, birth_place, birth_city, doc_type, doc_seria, doc_number, doc_date, doc_end)
        return redirect(url_for('comment'))


if __name__ == "__main__":
    app.run()