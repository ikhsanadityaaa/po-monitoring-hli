#!/bin/bash
set -e

echo "Initializing DB..."
python -c "
from app import app, db, _ensure_extra_columns
with app.app_context():
    db.create_all()
    _ensure_extra_columns()
    print('DB ready.')
"

echo "Starting gunicorn..."
exec gunicorn -w 1 -b :8080 --worker-tmp-dir /tmp --timeout 120 app:app
