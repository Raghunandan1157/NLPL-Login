#!/usr/bin/env python3
"""Live-reload dev server for NLPL Dashboard on port 8080."""

import livereload

server = livereload.Server()
root = "."

# Watch HTML, CSS, and JS for changes → auto-reload browser
server.watch("*.html")
server.watch("css/*.css")
server.watch("js/*.js")

print("Live server → http://localhost:8080")
server.serve(root=root, port=8080, open_url_delay=1)
