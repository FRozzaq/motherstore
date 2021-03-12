from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField
from wtforms.validators import DataRequired


# configure wtforms
class CreatePenjualanForm(FlaskForm):
    jml_penjualan = StringField("Penjualan")
