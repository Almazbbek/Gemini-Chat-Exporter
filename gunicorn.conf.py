import os

# Port is read from the environment here, so no shell expansion is needed
# in the start command.
bind = f"0.0.0.0:{os.environ.get('PORT', '8000')}"
worker_class = "gevent"
workers = 1
timeout = 300
max_requests = 50
max_requests_jitter = 10
