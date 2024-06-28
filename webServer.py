import http.server
import socketserver
import os

os.chdir(r"C:\Users\Kelvin.Su\connect")  

# Define the request handler to serve the file
class MyHttpRequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/':
            self.path = 'testFile.txt'
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.send_header('Refresh', '3') 
            self.end_headers()
            with open('testFile.txt', 'r') as file:
                lines = file.readlines()
                formattedLines = "<br>".join(lines)
                html = f"""
                <html>
                <head>
                    <style>
                    body {{
                        font-family: Arial, sans-serif;
                        color: #333;
                    }}
                    </style>
                </head>
                <body>
                {formattedLines}
                </body>
                </html>
                """
                self.wfile.write(html.encode())
        else:
            return http.server.SimpleHTTPRequestHandler.do_GET(self)

# Define the socket server
PORT = 8000
Handler = MyHttpRequestHandler

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    print("serving at port", PORT)
    httpd.serve_forever()
