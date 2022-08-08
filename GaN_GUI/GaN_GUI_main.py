from flask import Flask
from flaskwebgui import FlaskUI #get the FlaskUI class

DEBUG =True

app = Flask(__name__)

# Feed it the flask app instance 
ui = FlaskUI(app)

# do your logic as usual in Flask
@app.route("/")
def index():
    return "<h1> GaN Activity Guides generator</h1>"


if DEBUG:
    app.run()
else:
    # call the 'run' method
    ui.run()