from work_assistant.txtapp import app
from waitress import serve
import logging

# Setup Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('waitress')

if __name__ == "__main__":
    print("-------------------------------------------------------")
    print("  Administrative Assistant Production Server")
    print("  Status: Running")
    print("  Access URL: http://0.0.0.0:8080")
    print("  (Accessible from this machine at http://localhost:8080)")
    print("-------------------------------------------------------")
    
    # Run the server on port 8080
    serve(app, host='0.0.0.0', port=8080, threads=6)
