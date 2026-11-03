#!/bin/bash
cd "$(dirname "$0")" || exit 1
python3 -m venv .venv-mac || exit 1
.venv-mac/bin/python -m pip install -r requirements.txt || exit 1

# The Mac preview gets its own local port when 8765 is already taken.
# Nothing using the old port is stopped or contacted by this preview.
preview_port=$(.venv-mac/bin/python - <<'PY'
import socket
with socket.socket() as listener:
    try:
        listener.bind(('127.0.0.1', 8765))
    except OSError:
        listener.bind(('127.0.0.1', 0))
    print(listener.getsockname()[1])
PY
) || exit 1
export GLAZOK_PREVIEW_PORT="$preview_port"
echo "Глазок использует локальный порт $GLAZOK_PREVIEW_PORT"

# Both Tk windows share this Terminal. Close the client to stop this preview.
.venv-mac/bin/python server_app.py &
server_pid=$!
trap 'kill "$server_pid" 2>/dev/null || true' EXIT

# Wait until the server accepts connections before starting the client.
server_ready=0
for attempt in {1..30}; do
    if .venv-mac/bin/python -c "import json,os; from urllib.request import urlopen; url='http://127.0.0.1:'+os.environ['GLAZOK_PREVIEW_PORT']+'/health'; assert json.load(urlopen(url,timeout=1)).get('service')=='Глазок Сервер'" 2>/dev/null; then
        server_ready=1
        break
    fi
    sleep 0.2
done
if [ "$server_ready" -ne 1 ]; then
    echo "Сервер не запустился. Сообщение об ошибке смотрите выше."
    exit 1
fi
.venv-mac/bin/python glasok.py
