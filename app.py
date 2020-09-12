from flask import Flask, redirect, url_for, render_template, request, flash
import models as db_handler
import json
from flask import Response

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
        patronymic = request.form["patronymic"]
        city = request.form["city"]
        phone = request.form["phone"]
        mail = request.form["email"]
        comment = request.form["comment"]
        db_handler.add_new_comment(surname, name, patronymic, city, phone, mail, comment)
        return redirect(url_for('comment'))


if __name__ == "__main__":
    app.run()