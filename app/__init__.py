from flask import Flask
import os

def create_app():
    app = Flask(__name__)

    # Create required folders
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)

    from .routes import home, fetch
    app.register_blueprint(home.bp)
    app.register_blueprint(fetch.bp)

    return app
