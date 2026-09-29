# Local-development image only: the simulator is not a security boundary.
# Publish its port on loopback (-p 127.0.0.1:8787:8787), never on a network.
FROM python:3.12-slim

WORKDIR /sim
COPY server.py test_server.py ./

# Inside the container the loopback default would be unreachable through a
# published port, so listen on every container interface. SIM_BASE_URL is the
# URL the host reaches that port at; change it with the published host port.
ENV SIM_HOST=0.0.0.0 \
    SIM_PORT=8787 \
    SIM_BASE_URL=http://127.0.0.1:8787 \
    PYTHONUNBUFFERED=1

EXPOSE 8787

HEALTHCHECK --interval=5s --timeout=3s --retries=10 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8787/health', timeout=2)"]

CMD ["python", "server.py"]
