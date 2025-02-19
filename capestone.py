from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/email', methods=['POST'])
def categorize_email():
    data = request.json  # Get the JSON data sent by the client
    body = data.get('body', '').lower()

    categories = []
    print('backend has been called')
    if 'urgent' in body:
        categories.append('Urgent')
    elif 'meeting' in body:
        categories.append('Meeting')
    elif 'invoice' in body:
        categories.append('Invoice')
    else:
        categories.append('other')

    return jsonify(categories)  # Return the categories as a JSON response

@app.route('/category-change', methods=['POST'])
def category_change():
    # Get the JSON data sent by the client
    data = request.json

    body = data.get('body', '').lower()

    print('Received request data for /category-change:')

    # No return value (no response)
    return '', 204  # 204 No Content response

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=3001)
