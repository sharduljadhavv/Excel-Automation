from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_login import LoginManager

app = Flask(__name__)  
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///hdfc.db'
app.config['SECRET_KEY'] = 'abcdefghij123456'
db = SQLAlchemy(app)
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)

from .models import ProcessedFiles, User
from .forms import LoginForm, RegistrationForm
from flask_app import routes
from .template_tags import datetimeformat

