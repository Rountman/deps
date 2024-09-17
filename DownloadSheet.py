import os
import requests

# URL ke stažení souboru
url = "https://airportpardubice.sharepoint.com/:x:/g/provoz/ETHY5u8NPRBDqjYKITqzGwUBoHhm8TcXny7Ty9RMh2lBHQ?download=1"

# Cesta, kam chcete soubor uložit
local_filename = 'downloadedFlights.xlsx'

def download_file(url, local_filename):
    with requests.get(url, stream=True) as response:
        response.raise_for_status()  # Zkontroluje, zda je požadavek úspěšný
        with open(local_filename, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):  # Stahuje po částech
                file.write(chunk)
    print(f'Soubor byl úspěšně stažen jako {local_filename}')

def main():
    if os.path.exists(local_filename):
        os.remove(local_filename)  # Pokud soubor existuje, smaže ho
    
    download_file(url, local_filename)  # Stáhne a uloží nový soubor

    print('Soubor byl stažen a uložen.')

if __name__ == "__main__":
    main()
