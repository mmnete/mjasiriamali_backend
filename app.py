from flask import Flask, request, send_file
from flask_cors import CORS
from document_editor import add_banner_helper

app = Flask(__name__)
CORS(app, resources={"/add_banner": {"origins": "https://thriving-chebakia-ffd31c.netlify.app/"}}, supports_credentials=True)

@app.route('/add_banner', methods=['POST'])
def add_banner():
    new_doc = request.files['document']
    result_document_path = add_banner_helper(new_doc)
    
    # Create response
    response = send_file(result_document_path, as_attachment=True)

    # Set CORS headers
    response.headers['Access-Control-Allow-Origin'] = 'https://thriving-chebakia-ffd31c.netlify.app'
    response.headers['Access-Control-Allow-Methods'] = 'POST'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'

    return response

if __name__ == '__main__':
    app.run(debug=True)


