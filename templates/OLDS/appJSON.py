from flask import Flask, request, send_file, render_template, redirect, url_for
from weasyprint import HTML, CSS
from pypdf import PdfWriter
import io
import csv
import zipfile
import tempfile
import shutil
import os
from datetime import datetime
import threading
import uuid
import json

app = Flask(__name__, static_folder='static')

try:
    with open('testi_diploma.json', 'r', encoding='utf-8') as f:
        testi_statici_diploma = json.load(f)
except FileNotFoundError:
    testi_statici_diploma = {}
    print("ATTENZIONE: File 'testi_diploma.json' non trovato. I testi fissi non verranno caricati.")

temp_pdf_batches = {}

CLEANUP_DELAY_SECONDS = 3600  # 1 ora

def cleanup_batch_data(batch_id):
    """Funzione per pulire i dati di un batch dopo un certo ritardo."""
    batch_info = temp_pdf_batches.pop(batch_id, None)
    if batch_info:
        try:
            shutil.rmtree(batch_info['temp_dir'])
            print(f"Pulita directory temporanea per batch {batch_id}: {batch_info['temp_dir']}")
        except Exception as e:
            print(f"Errore durante la pulizia della directory temporanea {batch_info['temp_dir']} per batch {batch_id}: {e}")

@app.route('/', methods=['GET'])
def homepage():
    return render_template('upload.html')

@app.route('/upload-data', methods=['POST'])
def upload_data():
    if 'data_file' not in request.files:
        return 'Nessun file caricato', 400

    file = request.files['data_file']
    if file.filename == '':
        return 'Nessun file selezionato', 400

    if file:
        file_content = file.stream.read().decode('utf-8')
        students_data = parse_diploma_data(file_content)

        if not students_data:
            return 'Impossibile leggere i dati dal file. Controlla il formato o che contenga dati validi.', 400

        data_creazione = datetime.now()
        nome_cartella = data_creazione.strftime('%Y-%m-%d')
        
        log_entries = []
        
        batch_id = str(uuid.uuid4())
        current_batch_temp_dir = tempfile.mkdtemp()
        
        generated_pdf_filenames = [] 

        for i, student in enumerate(students_data):
            student_data_for_template = {
                key.lower(): value for key, value in student.items()
            }

            modulo_value = student_data_for_template.get('modulo', '').strip()
            
            # Elenco di tutti i PDF generati per questo studente (diploma e camicia)
            student_pdfs = []
            
            # --- INIZIO: Generazione del PDF del Diploma ---
            template_filename = ''
            if modulo_value == 'forml29v7':
                template_filename = 'diploma_forml29v7.html'
            elif modulo_value == 'forml00v7':
                template_filename = 'diploma_forml00v7.html'
            elif modulo_value == 'forml01v7':
                template_filename = 'diploma_forml01v7.html'
            else:
                log_entries.append(f"ATTENZIONE: Modulo '{modulo_value}' non riconosciuto per {student_data_for_template.get('nom_cog', 'uno studente')}. Saltando la generazione del PDF.")
                continue 

            student_data_for_template['lode'] = student_data_for_template.get('lode', '').upper().strip()
            student_data_for_template['testo_footer_fisso'] = "Imposta di bollo assolta in modo virtuale. Autorizzazione Intendenza di Finanza di Roma n.9120/88"
            firmar_val = student_data_for_template.get('firmar')
            student_data_for_template['firmar'] = f"{firmar_val}.png" if firmar_val else ""
            firmap_val = student_data_for_template.get('firmap')
            student_data_for_template['firmap'] = f"{firmap_val}.png" if firmap_val else ""
            firmad_val = student_data_for_template.get('firmad')
            student_data_for_template['firmad'] = f"{firmad_val}.png" if firmad_val else ""
            
            rendered_html = render_template(
                template_filename,
                testi_statici=testi_statici_diploma,
                **student_data_for_template
            )

            try:
                html_doc = HTML(string=rendered_html, base_url=request.url_root)
                pdf_bytes = html_doc.write_pdf()
                student_name_for_filename = student_data_for_template.get('nom_cog', f'studente_{i+1}').replace(' ', '_')
                pdf_filename = f'diploma_{student_name_for_filename}_{modulo_value}.pdf'
                pdf_path = os.path.join(current_batch_temp_dir, pdf_filename)
                with open(pdf_path, 'wb') as f:
                    f.write(pdf_bytes)
                
                student_pdfs.append(pdf_filename)
                log_entries.append(f"Diploma generato per: {student_data_for_template.get('nom_cog', 'N/A')}")

            except Exception as e:
                print(f"Errore nella generazione del PDF per {student_data_for_template.get('nom_cog', 'uno studente')}: {e}")
                log_entries.append(f"ERRORE: Impossibile generare il PDF per {student_data_for_template.get('nom_cog', 'uno studente')}. Errore: {e}")
            
            # --- FINE: Generazione del PDF del Diploma ---
            
            # --- INIZIO: Generazione del PDF della Camicia ---
            camicia_data_for_template = {
                'corso_laurea': student_data_for_template.get('corsolau', ''),
                'nome_studente': student_data_for_template.get('nom_cog', ''),
                'luogo_nascita': student_data_for_template.get('luogonas', '').split('(')[0].strip() if student_data_for_template.get('luogonas') and '(' in student_data_for_template.get('luogonas') else student_data_for_template.get('luogonas', '').strip(),
                'provincia_nascita': student_data_for_template.get('luogonas', '').split('(')[1].replace(')', '').strip() if '(' in student_data_for_template.get('luogonas', '') else '',
                'data_nascita': student_data_for_template.get('datanas', ''),
                'data_stampa': student_data_for_template.get('datastamp', ''),
                'numero_protocollo': student_data_for_template.get('protocol', ''),
                'data_rilascio': student_data_for_template.get('datastamp', ''),
                'numero_diploma': student_data_for_template.get('npergamena', ''),
                'genere_nato_nata': student_data_for_template.get('sesso', 'nato a').strip(), 
                'classe_laurea_dinamica': student_data_for_template.get('indicorso', ''),
                'firmad': student_data_for_template.get('firmad', ''),
                'firmar': student_data_for_template.get('firmar', ''),
                'firmap': student_data_for_template.get('firmap', '')  
            }

            rendered_camicia_html = render_template(
                'camicia_template.html',
                **camicia_data_for_template
            )
            
            try:
                html_camicia_doc = HTML(string=rendered_camicia_html, base_url=request.url_root)
                pdf_camicia_bytes = html_camicia_doc.write_pdf()

                camicia_pdf_filename = f'camicia_{student_name_for_filename}.pdf'
                camicia_pdf_path = os.path.join(current_batch_temp_dir, camicia_pdf_filename)
                with open(camicia_pdf_path, 'wb') as f:
                    f.write(pdf_camicia_bytes)
                
                student_pdfs.append(camicia_pdf_filename)
                log_entries.append(f"Camicia generata per: {student_data_for_template.get('nom_cog', 'N/A')}")

            except Exception as e:
                print(f"Errore nella generazione della Camicia per {student_data_for_template.get('nom_cog', 'uno studente')}: {e}")
                log_entries.append(f"ERRORE: Impossibile generare la Camicia per {student_data_for_template.get('nom_cog', 'uno studente')}. Errore: {e}")
            # --- FINE: Generazione del PDF della Camicia ---

            # Aggiungi i nomi dei file generati per questo studente alla lista principale
            generated_pdf_filenames.extend(student_pdfs)

        # --- INIZIO: Generazione del PDF combinato (solo diplomi) ---
        diploma_filenames = [f for f in generated_pdf_filenames if f.startswith('diploma_')]
        
        if diploma_filenames:
            merger = PdfWriter()
            combined_pdf_filename = f'tutti_i_diplomi_{nome_cartella}.pdf'
            combined_pdf_path = os.path.join(current_batch_temp_dir, combined_pdf_filename)
            for pdf_filename in diploma_filenames:
                file_path_to_merge = os.path.join(current_batch_temp_dir, pdf_filename)
                try:
                    merger.append(file_path_to_merge)
                except Exception as e:
                    print(f"ATTENZIONE: Impossibile unire il file {pdf_filename} al PDF combinato. Errore: {e}")
            merger.write(combined_pdf_path)
            merger.close()
            
            # Aggiungi il nome del file combinato alla lista per il download
            generated_pdf_filenames.append(combined_pdf_filename)
        # --- FINE: Generazione del PDF combinato ---
        
        log_content = '\n'.join(log_entries)
        
        log_file_path = os.path.join(current_batch_temp_dir, 'log_creazione_diplomi.txt')
        with open(log_file_path, 'w', encoding='utf-8') as f:
            f.write(log_content)
            
        temp_pdf_batches[batch_id] = {
            'temp_dir': current_batch_temp_dir,
            'filenames': generated_pdf_filenames,
            'log_content': log_content,
            'log_file_path': log_file_path,
            'original_folder_name': nome_cartella
        }
        
        timer = threading.Timer(CLEANUP_DELAY_SECONDS, cleanup_batch_data, args=[batch_id])
        timer.start()

        return redirect(url_for('preview_pdfs', batch_id=batch_id))

    return 'Errore sconosciuto', 500


#### Rotte di Servizio dei File


@app.route('/preview/<batch_id>')
def preview_pdfs(batch_id):
    batch_info = temp_pdf_batches.get(batch_id)
    if not batch_info:
        return "Anteprima non trovata o scaduta.", 404

    pdf_list_for_template = []
    # Prepara la lista di PDF per il template, includendo SOLO i diplomi
    for filename in batch_info['filenames']:
        # Aggiungi questa condizione per filtrare solo i PDF dei diplomi
        if filename.startswith('diploma_'): 
            pdf_list_for_template.append({
                'name': filename,
                'url': url_for('get_single_pdf', batch_id=batch_id, filename=filename)
            })

    return render_template('preview.html',
                            pdf_list=pdf_list_for_template,
                            download_url=url_for('download_zip_for_preview', batch_id=batch_id),
                            log_url=url_for('get_log_for_preview', batch_id=batch_id),
                            cleanup_delay_minutes=CLEANUP_DELAY_SECONDS / 60)

@app.route('/preview/pdf/<batch_id>/<filename>')
def get_single_pdf(batch_id, filename):
    batch_info = temp_pdf_batches.get(batch_id)
    if not batch_info:
        return "File non trovato.", 404
    
    if filename not in batch_info['filenames']:
        return "File non autorizzato o non trovato nel batch.", 403

    return send_file(os.path.join(batch_info['temp_dir'], filename), mimetype='application/pdf')

@app.route('/preview/log/<batch_id>')
def get_log_for_preview(batch_id):
    batch_info = temp_pdf_batches.get(batch_id)
    if not batch_info:
        return "Log non trovato o scaduto.", 404
    
    return send_file(batch_info['log_file_path'], 
                    mimetype='text/plain', 
                    as_attachment=True, 
                    download_name='log_creazione_diplomi.txt')


@app.route('/download_zip/<batch_id>')
def download_zip_for_preview(batch_id):
    batch_info = temp_pdf_batches.get(batch_id)
    if not batch_info:
        return "Download non trovato o scaduto.", 404

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zf:
        for filename in batch_info['filenames']:
            file_path = os.path.join(batch_info['temp_dir'], filename)
            
            if filename.startswith('diploma_'):
                subfolder = 'pergamene'
            elif filename.startswith('camicia_'):
                subfolder = 'camicie'
            elif filename.startswith('tutti_i_diplomi_'):
                subfolder = 'combinato'
            else:
                subfolder = 'altri' 

            arcname = os.path.join(batch_info['original_folder_name'], subfolder, filename)
            zf.write(file_path, arcname=arcname)
        
        log_filename_in_zip = os.path.join(batch_info['original_folder_name'], 'log_creazione_documenti.txt')
        zf.write(batch_info['log_file_path'], arcname=log_filename_in_zip)
    
    zip_buffer.seek(0)
    
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f'documenti_{batch_info["original_folder_name"]}.zip'
    )

def parse_diploma_data(file_content):
    lines = file_content.splitlines()
    if len(lines) < 4:
        return []

    header_line = lines[3]
    data_lines = lines[4:]

    reader = csv.reader(io.StringIO(header_line + '\n' + '\n'.join(data_lines)), delimiter='^')

    try:
        headers = next(reader)
    except StopIteration:
        return []

    students_data = []
    for row in reader:
        if row and len(row) == len(headers):
            student_dict = {}
            for i, header in enumerate(headers):
                student_dict[header.strip()] = row[i].strip()
            students_data.append(student_dict)
        else:
            pass

    return students_data

if __name__ == '__main__':
    app.run(debug=True)