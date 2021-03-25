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
        patranomic = request.form["patranomic"]

        exel_handler.add_new_exel(surname, name, patranomic)
        return redirect(url_for('comment'))


if __name__ == "__main__":
    app.run()
