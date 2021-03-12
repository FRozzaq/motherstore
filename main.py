from flask import Flask, redirect, render_template, request
from flask_bootstrap import Bootstrap
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import relationship
from flask_wtf import FlaskForm
from wtforms import IntegerField
import os
from newreport import NewReport
from fillreport import FillReport
from sqlalchemy import ForeignKey, Column, Integer, String


app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY")
Bootstrap(app)


@app.route('/')
def home():
    data = FillReport()
    vendor_names = data.get_vendor_name()
    print(vendor_names[1:])
    return render_template("index.html", vendors=vendor_names)


@app.route('/new')
def new_report():
    template = NewReport()
    template.new_report()
    return redirect('/')


@app.route('/fill/<vendor_name>', methods=['GET', 'POST'])
def fill_report(vendor_name):
    data = FillReport()
    wb = data.load_wb()
    wb_sheet = data.get_vendor_data(vendor_name)
    products = [product.value for product in wb_sheet["B"][12:]]
    stok_awal = [stok.value for stok in wb_sheet["F"][12:]]

    if request.method == "POST":
        vendor = wb[vendor_name]
        for i in range(len(products[:-2])):
            vendor[f"H{i+13}"] = int(request.form[f"product{i+1}"])
        wb.save(data.get_file_name())
        return redirect('/')

    return render_template("fill.html", name=vendor_name, data_count=len(products[:-2]), products=products, i=0, stok=stok_awal)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
