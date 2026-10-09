"""Mintresponse Work Orders — web application built on the Excel tracker."""
import os
import secrets

from flask import Flask


def create_app(test_config=None):
    app = Flask(__name__, instance_relative_config=True)
    os.makedirs(app.instance_path, exist_ok=True)
    app.config.update(
        DATABASE=os.environ.get("MINT_DATABASE",
                                os.path.join(app.instance_path, "mint.sqlite3")),
        TEMPLATE_XLSX=os.path.join(os.path.dirname(app.root_path), "data",
                                   "work_order_template.xlsx"),
        APP_PASSWORD=os.environ.get("MINT_PASSWORD") or None,
        MAX_CONTENT_LENGTH=20 * 1024 * 1024,
        SECRET_KEY=os.environ.get("MINT_SECRET_KEY") or _local_secret(app.instance_path),
    )
    if test_config:
        app.config.update(test_config)

    from . import db, views
    db.init_db(app.config["DATABASE"])
    app.teardown_appcontext(db.close_db)
    app.register_blueprint(views.bp)
    views.register_template_helpers(app)
    return app


def _local_secret(instance_path):
    """A random session key kept in the instance folder (not in git)."""
    path = os.path.join(instance_path, "secret_key")
    if not os.path.exists(path):
        with open(path, "w") as fh:
            fh.write(secrets.token_hex(32))
    with open(path) as fh:
        return fh.read().strip()
