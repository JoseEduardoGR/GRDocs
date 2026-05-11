import os
import sys
import yaml
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from colorama import Fore, init, Style

# Agregar directorios al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "doc"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "xlsx"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "pptx"))

from word import WordScriptGenerator
from excel import ExcelScriptGenerator
from powerpoint import PowerPointScriptGenerator

init(autoreset=True)

# Cargar configuración
with open("settings.yaml") as f:
    settings = yaml.safe_load(f)

app = Flask(__name__)
CORS(app)  # Permitir CORS para todas las rutas

# Generadores (se inicializan bajo demanda)
_word_gen = None
_excel_gen = None
_pptx_gen = None


def check_script_for_errors(script_path):
    """
    Verifica si el script generado contiene errores de API.
    Retorna (has_error, error_response) donde error_response es None si no hay error.
    """
    if not script_path or not os.path.exists(script_path):
        return False, None
    
    try:
        with open(script_path, 'r') as f:
            script_content = f.read()
        
        # Detectar rate limit
        if '[Error HTTP 429]' in script_content or 'rate-limited' in script_content.lower():
            return True, jsonify({
                "success": False,
                "error": "El modelo de IA está temporalmente limitado por tasa de uso",
                "error_type": "rate_limit",
                "suggestion": "Intenta de nuevo en unos segundos o cambia el modelo en settings.yaml",
                "retry_after": 12,
                "current_model": settings.get("model", "unknown")
            }), 429
        
        # Detectar otros errores de API
        if '[Error' in script_content and 'HTTP' in script_content:
            return True, jsonify({
                "success": False,
                "error": "Error al comunicarse con el modelo de IA",
                "error_type": "api_error",
                "suggestion": "Verifica tu API key o intenta con otro modelo",
                "current_model": settings.get("model", "unknown")
            }), 500
        
        return False, None
    except Exception:
        return False, None


def handle_runtime_error(e, dependency_name="Node.js"):
    """Maneja errores de runtime y retorna respuesta JSON apropiada."""
    error_msg = str(e)
    
    # Detectar error de sintaxis de ES6 modules (Node.js antiguo)
    if 'SyntaxError: Unexpected token' in error_msg and 'import' in error_msg:
        return jsonify({
            "success": False,
            "error": "Versión de Node.js incompatible (requiere v14+)",
            "error_type": "dependency_version_error",
            "details": "Tu versión de Node.js no soporta módulos ES6",
            "suggestion": "Actualiza Node.js a la versión 14 o superior: https://nodejs.org/",
            "current_nodejs_version": "< v14 (detectado por error de sintaxis)"
        }), 500
    
    if 'rate-limited' in error_msg.lower() or '429' in error_msg:
        return jsonify({
            "success": False,
            "error": "Modelo temporalmente limitado por tasa de uso",
            "error_type": "rate_limit",
            "details": error_msg,
            "suggestion": "Espera unos segundos o cambia el modelo en settings.yaml",
            "retry_after": 12,
            "current_model": settings.get("model", "unknown")
        }), 429
    elif dependency_name.lower() in error_msg.lower():
        return jsonify({
            "success": False,
            "error": f"{dependency_name} no está instalado o no está en el PATH",
            "error_type": "dependency_error",
            "suggestion": f"Instala {dependency_name}"
        }), 500
    else:
        return jsonify({
            "success": False,
            "error": error_msg,
            "error_type": "runtime_error"
        }), 500


def get_word_generator():
    """Obtener generador de Word (lazy loading)."""
    global _word_gen
    if _word_gen is None:
        _word_gen = WordScriptGenerator()  # Usa verbose desde settings.yaml
    return _word_gen


def get_excel_generator():
    """Obtener generador de Excel (lazy loading)."""
    global _excel_gen
    if _excel_gen is None:
        _excel_gen = ExcelScriptGenerator()  # Usa verbose desde settings.yaml
    return _excel_gen


def get_pptx_generator():
    """Obtener generador de PowerPoint (lazy loading)."""
    global _pptx_gen
    if _pptx_gen is None:
        _pptx_gen = PowerPointScriptGenerator()  # Usa verbose desde settings.yaml
    return _pptx_gen


@app.route('/', methods=['GET'])
def home():
    """Endpoint de bienvenida con documentación de la API."""
    return jsonify({
        "name": "GR Docs API",
        "version": "1.0.0",
        "description": "API para generar documentos Word, Excel y PowerPoint usando IA",
        "endpoints": {
            "POST /docx": "Generar documento Word (.docx)",
            "POST /xlsx": "Generar archivo Excel (.xlsx)",
            "POST /pptx": "Generar presentación PowerPoint (.pptx)",
            "GET /health": "Verificar estado del servidor"
        },
        "documentation": "https://github.com/JoseEduardoGR/GRDocs"
    })


@app.route('/health', methods=['GET'])
def health():
    """Verificar estado del servidor."""
    return jsonify({
        "status": "healthy",
        "model": settings.get("model", "unknown"),
        "generators": {
            "word": "ready",
            "excel": "ready",
            "powerpoint": "ready"
        }
    })


@app.route('/docx', methods=['POST'])
def generate_docx():
    """
    Generar documento Word (.docx)
    
    Body JSON:
    {
        "request": "Descripción del documento a generar",
        "download": true/false (opcional, default: false)
    }
    
    Response:
    {
        "success": true,
        "script_path": "ruta del script generado",
        "document_path": "ruta del documento generado",
        "message": "Documento generado exitosamente"
    }
    """
    try:
        data = request.get_json()
        
        if not data or 'request' not in data:
            return jsonify({
                "success": False,
                "error": "Campo 'request' es requerido",
                "error_type": "validation_error"
            }), 400
        
        user_request = data['request']
        should_download = data.get('download', False)
        
        print(Fore.CYAN + f"[API] Generando documento Word...")
        print(Fore.WHITE + f"Solicitud: {user_request[:100]}...")
        
        # Obtener generador
        word_gen = get_word_generator()
        
        # Generar y ejecutar
        script_path, docx_path = word_gen.generate_and_execute(user_request)
        
        # Verificar si hubo error en la generación del script
        has_error, error_response = check_script_for_errors(script_path)
        if has_error:
            print(Fore.RED + f"[API] ✗ Error detectado en el script generado")
            return error_response
        
        # Renombrar si se solicita descarga
        if should_download:
            custom_name = data.get('filename')
            final_path = word_gen.download(docx_path, custom_name)
        else:
            final_path = docx_path
        
        print(Fore.GREEN + f"[API] ✓ Documento generado: {final_path}")
        
        response = {
            "success": True,
            "script_path": script_path,
            "document_path": final_path,
            "message": "Documento Word generado exitosamente"
        }
        
        # Si se solicita descarga, enviar el archivo
        if should_download and data.get('send_file', False):
            return send_file(
                final_path,
                as_attachment=True,
                download_name=os.path.basename(final_path),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        
        return jsonify(response), 200
        
    except RuntimeError as e:
        print(Fore.RED + f"[API] ✗ RuntimeError: {str(e)}")
        return handle_runtime_error(e, "Node.js")
            
    except FileNotFoundError as e:
        print(Fore.RED + f"[API] ✗ FileNotFoundError: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e),
            "error_type": "file_not_found",
            "suggestion": "Verifica que todas las dependencias estén instaladas"
        }), 500
        
    except Exception as e:
        print(Fore.RED + f"[API] ✗ Error inesperado: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e),
            "error_type": "unknown_error"
        }), 500


@app.route('/xlsx', methods=['POST'])
def generate_xlsx():
    """
    Generar archivo Excel (.xlsx)
    
    Body JSON:
    {
        "request": "Descripción del archivo Excel a generar",
        "download": true/false (opcional, default: false)
    }
    
    Response:
    {
        "success": true,
        "script_path": "ruta del script generado",
        "file_path": "ruta del archivo generado",
        "message": "Archivo Excel generado exitosamente"
    }
    """
    try:
        data = request.get_json()
        
        if not data or 'request' not in data:
            return jsonify({
                "success": False,
                "error": "Campo 'request' es requerido",
                "error_type": "validation_error"
            }), 400
        
        user_request = data['request']
        should_download = data.get('download', False)
        
        print(Fore.CYAN + f"[API] Generando archivo Excel...")
        print(Fore.WHITE + f"Solicitud: {user_request[:100]}...")
        
        # Obtener generador
        excel_gen = get_excel_generator()
        
        # Generar y ejecutar
        script_path, xlsx_path = excel_gen.generate_and_execute(user_request)
        
        # Verificar si hubo error en la generación del script
        has_error, error_response = check_script_for_errors(script_path)
        if has_error:
            print(Fore.RED + f"[API] ✗ Error detectado en el script generado")
            return error_response
        
        # Renombrar si se solicita descarga
        if should_download:
            custom_name = data.get('filename')
            final_path = excel_gen.download(xlsx_path, custom_name)
        else:
            final_path = xlsx_path
        
        print(Fore.GREEN + f"[API] ✓ Archivo generado: {final_path}")
        
        response = {
            "success": True,
            "script_path": script_path,
            "file_path": final_path,
            "message": "Archivo Excel generado exitosamente"
        }
        
        # Si se solicita descarga, enviar el archivo
        if should_download and data.get('send_file', False):
            return send_file(
                final_path,
                as_attachment=True,
                download_name=os.path.basename(final_path),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        
        return jsonify(response), 200
        
    except RuntimeError as e:
        print(Fore.RED + f"[API] ✗ RuntimeError: {str(e)}")
        return handle_runtime_error(e, "Python")
            
    except FileNotFoundError as e:
        print(Fore.RED + f"[API] ✗ FileNotFoundError: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e),
            "error_type": "file_not_found",
            "suggestion": "Verifica que todas las dependencias estén instaladas"
        }), 500
        
    except Exception as e:
        print(Fore.RED + f"[API] ✗ Error inesperado: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e),
            "error_type": "unknown_error"
        }), 500


@app.route('/pptx', methods=['POST'])
def generate_pptx():
    """
    Generar presentación PowerPoint (.pptx)
    
    Body JSON:
    {
        "request": "Descripción de la presentación a generar",
        "download": true/false (opcional, default: false)
    }
    
    Response:
    {
        "success": true,
        "script_path": "ruta del script generado",
        "presentation_path": "ruta de la presentación generada",
        "message": "Presentación PowerPoint generada exitosamente"
    }
    """
    try:
        data = request.get_json()
        
        if not data or 'request' not in data:
            return jsonify({
                "success": False,
                "error": "Campo 'request' es requerido",
                "error_type": "validation_error"
            }), 400
        
        user_request = data['request']
        should_download = data.get('download', False)
        
        print(Fore.CYAN + f"[API] Generando presentación PowerPoint...")
        print(Fore.WHITE + f"Solicitud: {user_request[:100]}...")
        
        # Obtener generador
        pptx_gen = get_pptx_generator()
        
        # Generar y ejecutar
        script_path, pptx_path = pptx_gen.generate_and_execute(user_request)
        
        # Verificar si hubo error en la generación del script
        has_error, error_response = check_script_for_errors(script_path)
        if has_error:
            print(Fore.RED + f"[API] ✗ Error detectado en el script generado")
            return error_response
        
        # Renombrar si se solicita descarga
        if should_download:
            custom_name = data.get('filename')
            final_path = pptx_gen.download(pptx_path, custom_name)
        else:
            final_path = pptx_path
        
        print(Fore.GREEN + f"[API] ✓ Presentación generada: {final_path}")
        
        response = {
            "success": True,
            "script_path": script_path,
            "presentation_path": final_path,
            "message": "Presentación PowerPoint generada exitosamente"
        }
        
        # Si se solicita descarga, enviar el archivo
        if should_download and data.get('send_file', False):
            return send_file(
                final_path,
                as_attachment=True,
                download_name=os.path.basename(final_path),
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        
        return jsonify(response), 200
        
    except RuntimeError as e:
        print(Fore.RED + f"[API] ✗ RuntimeError: {str(e)}")
        return handle_runtime_error(e, "Node.js")
            
    except FileNotFoundError as e:
        print(Fore.RED + f"[API] ✗ FileNotFoundError: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e),
            "error_type": "file_not_found",
            "suggestion": "Verifica que todas las dependencias estén instaladas"
        }), 500
        
    except Exception as e:
        print(Fore.RED + f"[API] ✗ Error inesperado: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e),
            "error_type": "unknown_error"
        }), 500


@app.errorhandler(404)
def not_found(error):
    """Manejo de rutas no encontradas."""
    return jsonify({
        "success": False,
        "error": "Endpoint no encontrado",
        "available_endpoints": ["/", "/health", "/docx", "/xlsx", "/pptx"]
    }), 404


@app.errorhandler(500)
def internal_error(error):
    """Manejo de errores internos del servidor."""
    return jsonify({
        "success": False,
        "error": "Error interno del servidor"
    }), 500


def check_nodejs_version():
    """Verifica que Node.js esté instalado y sea versión 14+."""
    try:
        import subprocess
        result = subprocess.run(['node', '--version'], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            version_str = result.stdout.strip()
            # Extraer número de versión (ej: "v10.19.0" -> 10)
            version_major = int(version_str.lstrip('v').split('.')[0])
            
            if version_major < 14:
                print(Fore.RED + f"⚠️  ADVERTENCIA: Node.js {version_str} detectado")
                print(Fore.YELLOW + f"   GR Docs requiere Node.js v14+ para módulos ES6")
                print(Fore.YELLOW + f"   Actualiza desde: https://nodejs.org/\n")
                return False
            else:
                print(Fore.GREEN + f"✓ Node.js {version_str} detectado")
                return True
        else:
            print(Fore.RED + "⚠️  Node.js no encontrado en el PATH")
            return False
    except FileNotFoundError:
        print(Fore.RED + "⚠️  Node.js no está instalado")
        return False
    except Exception as e:
        print(Fore.YELLOW + f"⚠️  No se pudo verificar versión de Node.js: {e}")
        return False


def main():
    """Iniciar el servidor Flask."""
    port = settings.get('port', 8000)
    
    print(Fore.CYAN + Style.BRIGHT + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + Style.BRIGHT + "║                    GR DOCS API SERVER                          ║")
    print(Fore.CYAN + Style.BRIGHT + "╚════════════════════════════════════════════════════════════════╝\n")
    
    # Verificar Node.js
    check_nodejs_version()
    
    print(Fore.YELLOW + f"🚀 Servidor iniciando en http://localhost:{port}")
    print(Fore.GREEN + f"📝 Modelo configurado: {settings.get('model', 'unknown')}\n")
    
    print(Fore.WHITE + "Endpoints disponibles:")
    print(Fore.CYAN + f"  GET  http://localhost:{port}/")
    print(Fore.CYAN + f"  GET  http://localhost:{port}/health")
    print(Fore.GREEN + f"  POST http://localhost:{port}/docx")
    print(Fore.GREEN + f"  POST http://localhost:{port}/xlsx")
    print(Fore.GREEN + f"  POST http://localhost:{port}/pptx\n")
    
    print(Fore.YELLOW + "Presiona Ctrl+C para detener el servidor\n")
    
    app.run(
        host='0.0.0.0',
        port=port,
        debug=False
    )


if __name__ == '__main__':
    main()