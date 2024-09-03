from playwright.sync_api import sync_playwright, ElementHandle
from os.path import abspath, dirname, join
import pandas as pd


python_ubicacion = abspath(dirname(__file__))
ruta_pendiente = join(python_ubicacion, "pendiente.xlsx")

def leer_pendiente():
    df = pd.read_excel(ruta_pendiente)
    pendiente = df.to_dict("records")
    return pendiente

def cotizar(pendiente):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        link = "https://www.olvacourier.com/cotizador"
        page.goto(link)
        
        for row in pendiente:
            xpath = "//select[contains(@name, 'encuentras-departamento')]"
            page.select_option(xpath, row["origen_departamento"].capitalize())
            page.wait_for_timeout(1000)
            
            xpath = "//select[contains(@name, 'encuentras-provincia')]"
            page.select_option(xpath, row["origen_provincia"].capitalize())
            page.wait_for_timeout(1000)
            
            xpath = "//select[contains(@name, 'encuentras-distrito')]"
            page.select_option(xpath, row["origen_distrito"].capitalize())
            page.wait_for_timeout(1000)                        
            
            xpath = "//select[contains(@name, 'llevamos-departamento')]"
            page.select_option(xpath, row["destino_departamento"].capitalize())
            page.wait_for_timeout(1000)
            
            xpath = "//select[contains(@name, 'llevamos-provincia')]"
            page.select_option(xpath, row["destino_provincia"].capitalize())
            page.wait_for_timeout(1000)
            
            xpath = "//select[contains(@name, 'llevamos-distrito')]"
            page.select_option(xpath, row["destino_distrito"].capitalize())
            page.wait_for_timeout(1000)              
            
            xpath = "//strong[contains(text(), 'Paquetes')]"
            page.click(xpath)
            page.wait_for_timeout(1000)
            
            xpath = "//input[contains(@id, 'cotizador-pesa')]"
            page.fill(xpath, str(row["peso"]))
            page.wait_for_timeout(1000) 

            xpath = "//input[contains(@name, 'ancho')]"
            page.fill(xpath, str(row["ancho"]))
            page.wait_for_timeout(1000)

            xpath = "//input[contains(@name, 'largo')]"
            page.fill(xpath, str(row["largo"]))
            page.wait_for_timeout(1000)
            
            xpath = "//input[contains(@name, 'alto')]"
            page.fill(xpath, str(row["alto"]))
            page.wait_for_timeout(1000)
            
            xpath = "//strong[contains(text(), 'Cotizar')]"
            page.click(xpath)
            page.wait_for_timeout(1000)  
            
            xpath = "//b[@id='cotizador-estimado']"
            elem: ElementHandle = page.query_selector(xpath)
            dato = elem.text_content()
            page.wait_for_timeout(1000) 
            
            row["costo"] = dato
            
        page.close()
        browser.close()

def main():
    pendiente = leer_pendiente()
    cotizar(pendiente)
    
    df_respuesta = pd.DataFrame(pendiente)
    df_respuesta.to_excel(join(python_ubicacion, "respuesta.xlsx"), index=False)

if __name__ == "__main__":
    main()
