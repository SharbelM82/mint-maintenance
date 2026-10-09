import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mint import create_app  # noqa: E402


@pytest.fixture
def app(tmp_path):
    return create_app({"DATABASE": str(tmp_path / "test.sqlite3"), "TESTING": True,
                       "TESTING_NO_CSRF": True})


@pytest.fixture
def client(app):
    return app.test_client()
