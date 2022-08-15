from flask_wtf import FlaskForm
from wtforms import BooleanField,SubmitField, SelectMultipleField,SelectField,widgets, validators

years=[2020,2021,2022,2023,2024,2025]
north_consts = ["Select All", "Perseus", "Leo", "Bootes", "Cygnus", "Pegasus", "Orion", "Hercules"]
north_langs= ["Select All", "Catalan", "Chinese", "Czech", "English", "Finnish", "French", "Galician", "German", "Greek", "Indonesian", "Japanese", "Polish", "Portuguese", "Romanian", "Serbian", "Slovak", "Slovenian", "Spanish", "Swedish", "Thai"]
north_lats= ["Select All", "0", "10N", "20N", "30N", "40N", "50N"]

south_consts = ["Select All", "Orion","Canis Major", "Crux", "Leo", "Bootes", "Scorpius", "Hercules", "Sagittarius", "Grus", "Pegasus"]
south_langs = ["Select All", "English", "French", "Indonesian", "Portuguese", "Spanish"]
south_lats = ["Select All", "0", "10S", "20S", "30S", "40S"]



def select_multi_checkbox(field, ul_class='', **kwargs):
    kwargs.setdefault('type', 'checkbox')
    field_id = kwargs.pop('id', field.id)
    html = ['<ul %s>' % widgets.html_params(id=field_id, class_=ul_class)]
    for value, label, checked in field.iter_choices():
        choice_id = '%s-%s' % (field_id, value)
        options = dict(kwargs, name=field.name, value=value, id=choice_id)
        if checked:
            options['checked'] = 'checked'
        html.append('<li><input %s /> ' % widgets.html_params(**options))
        html.append('<label for="%s">%s</label></li>' % (field_id, label))
    html.append('</ul>')
    return ''.join(html)

class MyBooleanField(BooleanField):
    def __init__(self, label=None, validators=None, false_values=None, **kwargs):
        # don't accept blank as False, so that default will trigger
        super(MyBooleanField, self).__init__(label, validators, (False, 'false', 0, '0'), **kwargs)

    def process_formdata(self, valuelist):
        if not valuelist or valuelist[0] == '' or valuelist[0] is None:
            self.data = self.default
        elif valuelist[0] in self.false_values:
            self.data = False
        else:
            self.data = True

class SelectionsForm(FlaskForm):
    year=SelectField('Select Year',choices=(years))
    north_consts = SelectMultipleField('Constellations',choices=(north_consts), widget=select_multi_checkbox)
    north_langs = SelectMultipleField('Languagues',choices=(north_langs), widget=select_multi_checkbox)
    north_lats = SelectMultipleField('Latitudes',choices=(north_lats), widget=select_multi_checkbox)

    south_consts = SelectMultipleField('Constellations',choices=(south_consts), widget=select_multi_checkbox)
    south_langs = SelectMultipleField('Languagues',choices=(south_langs), widget=select_multi_checkbox)
    south_lats = SelectMultipleField('Latitudes',choices=(south_lats), widget=select_multi_checkbox)
    
    download_charts=MyBooleanField('Download Charts', validators=[validators.AnyOf([True, False])])
    select_everything=BooleanField('Select All')
    submit =SubmitField('Submit')

