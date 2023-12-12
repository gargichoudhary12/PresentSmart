import os
from flask import Flask, render_template, url_for, flash, redirect, request, send_from_directory, abort, jsonify
from flask_login import LoginManager, login_user, login_required, logout_user, current_user

from flask_bcrypt import Bcrypt
from database import db
from utils.content import chat_development
from utils.presentation import parse_response, create_ppt
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///site.db'
bcrypt = Bcrypt(app)
db.init_app(app)







@app.route("/")
@app.route("/home")
def home():
    return render_template('home.html', user=current_user)





@app.route('/generator', methods=['GET', 'POST'])
def generate():
    if request.method == 'POST':
        number_of_slide = request.form.get('number_of_slide')
        user_text = request.form.get('user_text')
        template_choice = request.form.get('template_choice')
        presentation_title = request.form.get('presentation_title')
        presenter_name = request.form.get('presenter_name')
        insert_image = 'insert_image' in request.form

        user_message = f"I want you to generate PowerPoint presentations. The number of slides is {number_of_slide}. " \
                       f"The content is: {user_text}.The title of content for each slide must be unique, " \
                       f"and extract the most important keyword within two words for each slide. Summarize the content for each slide. "

        assistant_response = chat_development(user_message)
        print(f"Assistant Response:\n{assistant_response}")
        slides_content = parse_response(assistant_response)
        create_ppt(slides_content, template_choice, presentation_title, presenter_name, insert_image)

    return render_template('generator.html', title='Generate')


@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        return send_from_directory('generated', filename, as_attachment=True)

    except FileNotFoundError:
        abort(404)


if __name__ == "__main__":

    app.run(port=5001, debug=True)