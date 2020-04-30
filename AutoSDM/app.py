import sdm
import incidente
import jsonpickle
import json
import resultados as rs
from flask import Flask, request, render_template, jsonify
from flask_request_params import bind_request_params

app = Flask(__name__)
# bind rails like params to request.params
app.before_request(bind_request_params)

def transform(object):
    return jsonpickle.encode(object)

@app.route('/exporta',methods=['GET','POST'])
def home():
    query = request.args.get('query')
    try:
        resultado = rs.buscar(query)
        resultado =  json.dumps(resultado, default = lambda x: x.__dict__,indent=True, sort_keys=False)
        return resultado
    except:
        return Exception


@app.route('/exporta/all',methods=['GET','POST'])
def exportall():
    try:
        resultado = rs.buscar(query)
        resultado =  json.dumps(resultado, default = lambda x: x.__dict__,indent=True, sort_keys=False)
        return resultado
    except:
        return Exception

@app.route('/echo/<path>', methods=['GET', 'POST'])
def echo(path):
    return jsonify(request.params)

@app.route('/user', methods=['POST'])
def create_user():
    user = request.params.require('user').permit('name', 'password')
    print(user)
    return jsonify(user)

app.run(port=5001)
app.run(debug=True)
