from flask import Flask, render_template, request
import random, threading, webbrowser
from fuzzywuzzy import fuzz, process
import numpy as np
import pandas as pd

app = Flask(__name__)
variable_dict = {}

@app.route('/')
def home():
    df = pd.DataFrame()
    return render_template('home.html',tables=[df.to_html(classes='data')], titles=df.columns.values)

def get_ratio(row):
    row1 = row['FirstName']
    name1=variable_dict['name1']
    return fuzz.token_sort_ratio(row1, name1)



@app.route('/result', methods=['POST', 'GET'])
def result():
    form_data = request.form
    print(form_data['value'])
    name1=form_data['value']

    filename=form_data['filename']
    variable_dict['name1'] = name1

    df = pd.read_excel('C:/Users/Dell laptop/Desktop/Rubix project string matching/HRM.xlsx')
    x=df[df.apply(get_ratio, axis=1) > 50]
    x = x.reset_index().drop(['index'], 1)
    path=r'C:\Users\Dell laptop\Desktop\Rubix project string matching'+'\\'+filename+'.xlsx'
    x.to_excel(path, index=False)
    return render_template('home.html',tables=[x.to_html(classes='data')], titles=x.columns.values)


if __name__ == "__main__":
    port = 5000
    url = "http://127.0.0.1:{0}".format(port)
    threading.Timer(1.25, lambda: webbrowser.open(url)).start()
    app.run(port=port, debug=False)