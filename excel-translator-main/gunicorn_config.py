# gunicorn_config.py
bind = "0.0.0.0:5000"
workers = 1
timeout = 300          # 5 分钟
keepalive = 2
preload_app = True
