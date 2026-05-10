import os
import sys
from colorama import init, Fore, Style

init(autoreset=True)

# Agregar directorios al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "doc"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "xlsx"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "pptx"))

from word import WordScriptGenerator
from excel import ExcelScriptGenerator
from powerpoint import PowerPointScriptGenerator


def test_word():
    """Ejemplo de generación de documento Word."""
    
    print(Fore.CYAN + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + "║          GENERADOR DE DOCUMENTOS WORD                          ║")
    print(Fore.CYAN + "╚════════════════════════════════════════════════════════════════╝\n")
    
    generator = WordScriptGenerator()
    print(Fore.YELLOW + "Generador Word inicializado\n")
    
    user_request = """
    Genera un documento de investigación académica sobre el instrumento musical "Violín".
    
    El documento debe incluir:
    - Portada con título "El Violín: Historia, Construcción y Técnica"
    - Índice de contenidos
    - Introducción sobre el violín
    - Historia y evolución del instrumento (desde sus orígenes hasta la actualidad)
    - Anatomía y construcción del violín (partes, materiales, proceso de fabricación)
    - Técnicas de interpretación (postura, arco, digitación)
    - Grandes maestros y compositores (Stradivari, Paganini, etc.)
    - Tabla comparativa de tipos de violines (barroco, clásico, moderno)
    - El violín en diferentes géneros musicales
    - Conclusiones
    - Bibliografía
    
    Usa colores elegantes: tonos burdeos y dorado para un estilo clásico musical.
    El documento debe tener un tono académico y profesional.
    """
    
    print(Fore.CYAN + "Solicitud:")
    print(Fore.WHITE + "Documento de investigación sobre el Violín\n")
    
    try:
        print(Fore.YELLOW + "[1/3] Generando script JavaScript...")
        script_path = generator.generate_script(user_request)
        print(Fore.GREEN + f"✓ Script generado: {script_path}\n")
        
        print(Fore.YELLOW + "[2/3] Ejecutando script con Node.js...")
        print(Fore.RED + "NOTA: Asegúrate de tener instalada la librería 'docx': npm install docx\n")
        
        docx_path = generator.execute_script(script_path)
        print(Fore.GREEN + f"✓ Documento generado: {docx_path}\n")
        
        print(Fore.YELLOW + "[3/3] Renombrando documento...")
        final_path = generator.download(docx_path)
        print(Fore.GREEN + f"✓ Documento final: {final_path}\n")
        
        print(Fore.CYAN + Style.BRIGHT + "¡Word generado exitosamente!")
        return final_path
        
    except Exception as e:
        print(Fore.RED + f"\n✗ Error: {e}")
        return None


def test_excel():
    """Ejemplo de generación de archivo Excel."""
    
    print(Fore.CYAN + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + "║          GENERADOR DE ARCHIVOS EXCEL                           ║")
    print(Fore.CYAN + "╚════════════════════════════════════════════════════════════════╝\n")
    
    generator = ExcelScriptGenerator()
    print(Fore.YELLOW + "Generador Excel inicializado\n")
    
    user_request = """
    Genera un archivo Excel profesional de "Control de Inventario de Instrumentos Musicales".
    
    El archivo debe incluir:
    - Hoja 1: "Inventario" con columnas:
      * ID del instrumento
      * Nombre del instrumento
      * Categoría (Cuerda, Viento, Percusión)
      * Marca
      * Modelo
      * Año de fabricación
      * Estado (Nuevo, Usado, Reparación)
      * Precio de compra
      * Precio de venta
      * Stock disponible
      * Ubicación en almacén
    
    - Hoja 2: "Resumen" con:
      * Total de instrumentos por categoría (tabla con gráfico)
      * Valor total del inventario
      * Instrumentos que necesitan reposición (stock < 5)
      * Top 5 instrumentos más caros
    
    - Hoja 3: "Ventas Mensuales" con:
      * Tabla de ventas por mes (últimos 12 meses)
      * Gráfico de línea de tendencia
      * Cálculo de promedio mensual
    
    Usa colores: azul marino para encabezados, verde para valores positivos.
    Incluye al menos 20 instrumentos de ejemplo con datos realistas.
    Todas las fórmulas deben ser de Excel, no valores hardcodeados.
    """
    
    print(Fore.CYAN + "Solicitud:")
    print(Fore.WHITE + "Control de Inventario de Instrumentos Musicales\n")
    
    try:
        print(Fore.YELLOW + "[1/3] Generando script Python...")
        script_path = generator.generate_script(user_request)
        print(Fore.GREEN + f"✓ Script generado: {script_path}\n")
        
        print(Fore.YELLOW + "[2/3] Ejecutando script con Python...")
        print(Fore.RED + "NOTA: Asegúrate de tener instalada la librería 'openpyxl': pip install openpyxl\n")
        
        xlsx_path = generator.execute_script(script_path)
        print(Fore.GREEN + f"✓ Archivo generado: {xlsx_path}\n")
        
        print(Fore.YELLOW + "[3/3] Renombrando archivo...")
        final_path = generator.download(xlsx_path)
        print(Fore.GREEN + f"✓ Archivo final: {final_path}\n")
        
        print(Fore.CYAN + Style.BRIGHT + "¡Excel generado exitosamente!")
        return final_path
        
    except Exception as e:
        print(Fore.RED + f"\n✗ Error: {e}")
        return None


def test_powerpoint():
    """Ejemplo de generación de presentación PowerPoint."""
    
    print(Fore.CYAN + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + "║          GENERADOR DE PRESENTACIONES POWERPOINT                ║")
    print(Fore.CYAN + "╚════════════════════════════════════════════════════════════════╝\n")
    
    generator = PowerPointScriptGenerator()
    print(Fore.YELLOW + "Generador PowerPoint inicializado\n")
    
    user_request = """
    Genera una presentación profesional sobre "Historia del Violín: De Stradivari a la Actualidad".
    
    La presentación debe tener aproximadamente 10-12 slides e incluir:
    
    - Slide 1: Portada elegante con título principal y subtítulo
    - Slide 2: Índice de contenidos
    - Slide 3: Introducción - ¿Qué es el violín?
    - Slide 4: Orígenes históricos (siglo XVI)
    - Slide 5: La era dorada - Antonio Stradivari
    - Slide 6: Anatomía del violín (diagrama con partes principales)
    - Slide 7: Técnicas de construcción tradicionales
    - Slide 8: Grandes violinistas de la historia (timeline)
    - Slide 9: El violín en diferentes géneros musicales (tabla comparativa)
    - Slide 10: El violín moderno vs. el violín barroco (comparación lado a lado)
    - Slide 11: Curiosidades y datos interesantes
    - Slide 12: Slide de cierre con mensaje final
    
    Usa colores elegantes: burdeos oscuro (#6B0F1A) y dorado (#D4AF37) para un estilo clásico musical.
    Incluye al menos una tabla comparativa y un timeline visual.
    Cada slide debe tener un layout diferente para mantener el interés visual.
    Tono: Académico pero accesible, con datos históricos precisos.
    """
    
    print(Fore.CYAN + "Solicitud:")
    print(Fore.WHITE + "Presentación sobre Historia del Violín\n")
    
    try:
        print(Fore.YELLOW + "[1/3] Generando script JavaScript...")
        script_path = generator.generate_script(user_request)
        print(Fore.GREEN + f"✓ Script generado: {script_path}\n")
        
        print(Fore.YELLOW + "[2/3] Ejecutando script con Node.js...")
        print(Fore.RED + "NOTA: Asegúrate de tener instalada la librería 'pptxgenjs': npm install pptxgenjs\n")
        
        pptx_path = generator.execute_script(script_path)
        print(Fore.GREEN + f"✓ Presentación generada: {pptx_path}\n")
        
        print(Fore.YELLOW + "[3/3] Renombrando presentación...")
        final_path = generator.download(pptx_path)
        print(Fore.GREEN + f"✓ Presentación final: {final_path}\n")
        
        print(Fore.CYAN + Style.BRIGHT + "¡PowerPoint generado exitosamente!")
        return final_path
        
    except Exception as e:
        print(Fore.RED + f"\n✗ Error: {e}")
        return None


def main():
    """Menú principal para elegir qué generar."""
    
    print(Fore.CYAN + Style.BRIGHT + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + Style.BRIGHT + "║              GR DOCS - GENERADOR DE DOCUMENTOS                 ║")
    print(Fore.CYAN + Style.BRIGHT + "╚════════════════════════════════════════════════════════════════╝\n")
    
    print(Fore.YELLOW + "Selecciona qué deseas generar:\n")
    print(Fore.WHITE + "  1. Documento Word (.docx)")
    print(Fore.WHITE + "  2. Archivo Excel (.xlsx)")
    print(Fore.WHITE + "  3. Presentación PowerPoint (.pptx)")
    print(Fore.WHITE + "  4. Todos\n")
    
    choice = input(Fore.CYAN + "Opción (1/2/3/4): " + Style.RESET_ALL).strip()
    
    results = []
    
    if choice == "1":
        result = test_word()
        if result:
            results.append(("Word", result))
    
    elif choice == "2":
        result = test_excel()
        if result:
            results.append(("Excel", result))
    
    elif choice == "3":
        result = test_powerpoint()
        if result:
            results.append(("PowerPoint", result))
    
    elif choice == "4":
        word_result = test_word()
        if word_result:
            results.append(("Word", word_result))
        
        excel_result = test_excel()
        if excel_result:
            results.append(("Excel", excel_result))
        
        pptx_result = test_powerpoint()
        if pptx_result:
            results.append(("PowerPoint", pptx_result))
    
    else:
        print(Fore.RED + "\n✗ Opción inválida")
        return
    
    # Resumen final
    if results:
        print(Fore.CYAN + Style.BRIGHT + "\n╔════════════════════════════════════════════════════════════════╗")
        print(Fore.CYAN + Style.BRIGHT + "║                    RESUMEN DE ARCHIVOS                         ║")
        print(Fore.CYAN + Style.BRIGHT + "╚════════════════════════════════════════════════════════════════╝\n")
        
        for doc_type, path in results:
            print(Fore.GREEN + f"✓ {doc_type}: " + Fore.WHITE + path)
        
        print()


if __name__ == "__main__":
    main()
