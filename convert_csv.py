import os
import pandas as pd
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        rawdata = f.read()
    result = chardet.detect(rawdata)
    return result['encoding']

def convert_csv_to_utf8_bom(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"Fehler: Die Datei '{input_file}' wurde nicht gefunden.")
        return
    
    try:
        # Encoding automatisch erkennen
        detected_encoding = detect_encoding(input_file)
        print(f"Erkanntes Encoding: {detected_encoding}")
        
        # Datei einlesen mit erkanntem Encoding
        df = pd.read_csv(input_file, encoding=detected_encoding)
        
        # Datei mit UTF-8 BOM speichern
        df.to_csv(output_file, encoding='utf-8-sig', index=False)
        
        print(f"Erfolgreich konvertiert: '{output_file}'")
    except Exception as e:
        print(f"Fehler bei der Verarbeitung: {e}")

if __name__ == "__main__":
    input_file = "System und Update Liste.temp.csv"
    output_file = "System und Update Liste.csv"
    
    convert_csv_to_utf8_bom(input_file, output_file)
