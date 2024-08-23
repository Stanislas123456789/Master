import os
import locale
import re 
from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from datetime import date, datetime
from translate import Translator
from jinja2 import Environment
import math
import webbrowser
import threading

# Set up the French locale
locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8') 

def format_thousand(number):
    """Format number with a thousand separator."""
    ### debug
    print(f"format_thousand called with value: {number} of type: {type(number)}")
    ### end debug
    if not number:
        return "0"
    number_str = str(number)
    try:
        # Attempt to remove non-breaking space and convert to int
        number = int(float(number_str.replace('\xa0', '')))
    except ValueError:
        # In case of any conversion error, log the issue and return 0 formatted
        print(f"Error formatting number: {number_str}")
        return locale.format_string("%d", 0, grouping=True)
    return locale.format_string("%d", number, grouping=True)

app = Flask(__name__, static_url_path='/static', template_folder=os.path.join(os.path.dirname(__file__), 'templates'))

@app.route('/', methods=['GET', 'POST'])

def formulaire():
    if request.method == 'POST':
        # Get form values
        seller = request.form['seller']
        seller_nickname = request.form['seller_nickname']
        phase_appel_offre = request.form['phase_appel_offre']
        type_creances = request.form['type_creances']
        transaction = request.form['transaction']
        debiteurs = request.form['debiteurs']
        mode_acquisition = request.form['mode_acquisition']
        presence_titres = request.form['presence_titres']
        banque_france = request.form['banque_france']
        mark_up_contrat = request.form['mark_up_contrat']
        nb_lots_stock = int(request.form['nb_lots_stock'])
        nb_lots_flux = int(request.form['nb_lots_flux'])

        # Check if Seller or Seller's Nickname fields are empty
        if not seller or not seller_nickname:
            error_message = "Please fill in the required fields."
            return render_template('formulaire.html', error_message=error_message)
        
        # Create dateFR (today's date in French)
        today_date = str(date.today().strftime("%d %B %Y"))
        translator = Translator(to_lang="FR")
        dateFR = translator.translate(text = today_date)

        # Reformat phase_appel_offre
        if phase_appel_offre == "BO":
            phase_appel_offre = "Engageante"
        else:
            phase_appel_offre = "Indicative"

        # Create a dictionary to store the form values
        data = {
            'seller': seller,
            'seller_nickname': seller_nickname,
            'type_offer': phase_appel_offre,
            'type_creances': type_creances,
            'transaction': transaction,
            'debiteurs': debiteurs,
            'buyer': mode_acquisition,
            'presence_titres': presence_titres,
            'banque_france': banque_france,
            'mark_up_contrat': mark_up_contrat,
            'nb_lots_stock': nb_lots_stock,
            'nb_lots_flux': nb_lots_flux,
            'stock_lots': {},
            'flux_lots': {},
            'today' : dateFR
        }

        def round_like_excel(number, ndigits=0):
            if ndigits < 0:
                raise ValueError("ndigits must be non-negative")
            
            multiplier = 10 ** ndigits
            if number >= 0:
                return math.floor(number * multiplier + 0.5) / multiplier
            else:
                return math.ceil(number * multiplier - 0.5) / multiplier

        def process_lot_data(lot_type, index):
            lot_description = request.form.get(f"lot_description_{lot_type}_{index}", "")
            lot_number = request.form.get(f"lot_number_{lot_type}_{index}", "")
            lot_amount = request.form.get(f"lot_amount_{lot_type}_{index}", "")
            lot_price = request.form.get(f"lot_price_{lot_type}_{index}", "")

            # Ensure we don't divide by zero or empty values
            average = int(float(lot_amount) / float(lot_number)) if lot_number and float(lot_number) != 0 else 0

            # Compute PPQ
            ppq = (float(lot_price) / float(lot_amount)) * 100 if float(lot_amount) != 0 else 0

            # Convert the computed PPQ to 2 decimal places
            ppq = round_like_excel(ppq, 2)

            ### debug
            print(f"Processing flux lot {i}, lot_amount: {lot_amount}, type: {type(lot_amount)}")
            ### end debug
            
            return {
                'Number': lot_number,
                'FV': format_thousand(lot_amount),
                'Price': format_thousand(lot_price),
                'Average': format_thousand(average),
                'PPQ' : ppq
            }
        
        # Populate stock lot details
        if nb_lots_stock > 0:
            for i in range(1, nb_lots_stock + 1):
                data['stock_lots'][f"{i}"] = process_lot_data('stock', i)
        
        # Populate flux lot details
        if nb_lots_flux > 0:
            for i in range(1, nb_lots_flux + 1):
                data['flux_lots'][f"{i}"] = process_lot_data('flux', i)
                
        # Calculate total_flux, total_stock, total_flux_FV, total_stock_FV, total_flux_price, and total_stock_price
        total_flux = sum(int(lot['Number']) for lot in data['flux_lots'].values())
        total_stock = sum(int(lot['Number']) for lot in data['stock_lots'].values())
        total_flux_FV = sum(int(lot['FV'].replace('\xa0', '')) for lot in data['flux_lots'].values())
        total_stock_FV = sum(int(lot['FV'].replace('\xa0', '')) for lot in data['stock_lots'].values())
        total_flux_price = sum(int(lot['Price'].replace('\xa0', '')) for lot in data['flux_lots'].values())
        total_stock_price = sum(int(lot['Price'].replace('\xa0', '')) for lot in data['stock_lots'].values())
        
        # Calculate total_flux_average and total_stock_average
        total_flux_average = int(round_like_excel(total_flux_FV / total_flux)) if total_flux != 0 else 0
        total_stock_average = int(round_like_excel(total_stock_FV / total_stock)) if total_stock != 0 else 0
        
        # Calculate total_stock_PPQ and total_flux_PPQ
        total_stock_PPQ = round_like_excel((total_stock_price / total_stock_FV) * 100, 2) if total_stock_FV != 0 else 0
        total_flux_PPQ = round_like_excel((total_flux_price / total_flux_FV) * 100, 2) if total_flux_FV != 0 else 0
        
        # Add the total values, averages, and PPQ values to the data dictionary
        data['total_flux'] = format_thousand(total_flux)
        data['total_stock'] = format_thousand(total_stock)
        data['total_flux_FV'] = format_thousand(total_flux_FV)
        data['total_stock_FV'] = format_thousand(total_stock_FV)
        data['total_flux_Price'] = format_thousand(total_flux_price)
        data['total_stock_Price'] = format_thousand(total_stock_price)
        data['total_flux_Average'] = format_thousand(total_flux_average)
        data['total_stock_Average'] = format_thousand(total_stock_average)
        data['total_stock_PPQ'] = total_stock_PPQ
        data['total_flux_PPQ'] = total_flux_PPQ

        # Print the data dictionary
        def print_data(data):
            for key, value in data.items():
                if isinstance(value, list):
                    print(f"{key}:")
                    for item in value:
                        print(f"  - {item}")
                else:
                    print(f"{key}: {value}")
        
        print_data(data)

        print(data)

        try:
            # Load the Word template using the custom class
            template = DocxTemplate('Template Offre.docx')
        except Exception as e:
            return f"Error loading template: {e}"

       # Render the template with the custom Jinja environment and form data
        template.render(data)

        # Save the rendered template to a new Word document
        output_file = "./offres/Offre_" + data['type_offer'] + "_" \
        + data['seller_nickname'] + "_EOSFR_" + data['today'] + "_" \
        + datetime.now().strftime("%H%M%S") + ".docx" 
        template.save(output_file)

        # Return the generated Word file for download
        return send_file(output_file, as_attachment=True, download_name=os.path.basename(output_file))
    
    return render_template('formulaire.html')

def open_browser():
      """Wait for the server to start and then open the browser."""
      print("Opening browser...")
      webbrowser.open_new('http://127.0.0.1:5000/')

if __name__ == '__main__':
    print("Starting Flask app...")
    threading.Timer(1.25, open_browser).start()
    app.run(debug=True, use_reloader=False)