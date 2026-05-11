import os
import sys
import uuid
import yaml
import subprocess
from pathlib import Path

# Agregar el directorio padre al path para importar desde assets
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from assets.engine import OpenRouterEngine


class ExcelScriptGenerator:
    """Genera scripts Python para crear archivos Excel usando el modelo de IA."""
    
    def __init__(self, settings_path: str = "settings.yaml", verbose: bool = None):
        self.settings = self._load_settings(settings_path)
        
        # Si verbose no se especifica, leer desde settings.yaml
        if verbose is None:
            verbose = self.settings.get("verbose", True)
        
        self.engine = OpenRouterEngine(settings_path, verbose=verbose)
        self.prompt_path = self.settings.get("prompt_xlsx", "xlsx/prompt.gr")
        self.cache_dir = Path("xlsx/cache")
        self.output_dir = self.cache_dir / "output"
        self._expected_output = None  # Para rastrear el output esperado
        
        # Crear directorios si no existen
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def _load_settings(self, path: str) -> dict:
        """Carga la configuración desde el archivo YAML."""
        try:
            with open(path) as f:
                return yaml.safe_load(f) or {}
        except FileNotFoundError:
            print(f"[ExcelScriptGenerator] settings.yaml no encontrado en '{path}'")
            return {}
    
    def _load_prompt(self) -> str:
        """Carga el prompt desde el archivo especificado en settings."""
        try:
            with open(self.prompt_path) as f:
                return f.read().strip()
        except FileNotFoundError:
            print(f"[ExcelScriptGenerator] Prompt no encontrado en '{self.prompt_path}'")
            return ""
    
    def generate_script(self, user_request: str) -> str:
        """Genera un script Python basado en el prompt y la solicitud del usuario.
        
        Args:
            user_request: Descripción del archivo Excel que el usuario quiere generar
            
        Returns:
            Ruta del script generado
        """
        # Cargar el prompt base
        base_prompt = self._load_prompt()
        
        # Generar nombre único para el output
        output_id = str(uuid.uuid4())
        output_filename = f"{output_id}.xlsx"
        output_path = self.output_dir / output_filename
        
        # Construir el prompt completo con la ruta de salida específica
        full_prompt = f"""{base_prompt}

Solicitud del usuario:
{user_request}

IMPORTANTE: El archivo de salida DEBE guardarse en la ruta: {output_path}

Genera SOLO el código Python completo, sin explicaciones adicionales.
Al final del script, asegúrate de incluir un print() indicando que el archivo se generó exitosamente con su tamaño."""
        
        # Generar el script usando el engine
        script_content = self.engine.process(full_prompt)
        
        # Limpiar el contenido (remover markdown si existe)
        script_content = self._clean_script(script_content)
        
        # Generar nombre único para el script
        script_id = str(uuid.uuid4())
        script_filename = f"{script_id}.py"
        script_path = self.cache_dir / script_filename
        
        # Guardar el script
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        # Guardar también la ruta del output esperado para referencia
        self._expected_output = str(output_path)
        
        print(f"[ExcelScriptGenerator] Script generado: {script_path}")
        print(f"[ExcelScriptGenerator] Output esperado: {output_path}")
        return str(script_path)
    
    def _clean_script(self, content: str) -> str:
        """Limpia el contenido del script removiendo markdown y texto extra."""
        lines = content.split('\n')
        cleaned_lines = []
        in_code_block = False
        
        for line in lines:
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
                continue
            if in_code_block or not line.strip().startswith('```'):
                cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    def execute_script(self, script_path: str) -> str:
        """Ejecuta el script Python y retorna la ruta del archivo Excel generado.
        
        Args:
            script_path: Ruta del script a ejecutar
            
        Returns:
            Ruta del archivo .xlsx generado
        """
        try:
            # Convertir a ruta absoluta
            abs_script_path = os.path.abspath(script_path)
            
            # Ejecutar el script con Python desde el directorio raíz del proyecto
            result = subprocess.run(
                [sys.executable, abs_script_path],
                capture_output=True,
                text=True,
                timeout=60,
                cwd=os.getcwd()
            )
            
            if result.returncode != 0:
                print(f"[ExcelScriptGenerator] Error al ejecutar script:")
                print(result.stderr)
                raise RuntimeError(f"Script execution failed: {result.stderr}")
            
            print(f"[ExcelScriptGenerator] Script ejecutado exitosamente")
            print(result.stdout)
            
            # Si tenemos el output esperado, verificar que existe
            if self._expected_output and os.path.exists(self._expected_output):
                return self._expected_output
            
            # Buscar el archivo .xlsx generado en el directorio de salida
            xlsx_files = list(self.output_dir.glob("*.xlsx"))
            
            if not xlsx_files:
                raise FileNotFoundError("No se encontró el archivo .xlsx generado")
            
            # Retornar el más reciente
            latest_xlsx = max(xlsx_files, key=lambda p: p.stat().st_mtime)
            return str(latest_xlsx)
            
        except subprocess.TimeoutExpired:
            raise RuntimeError("Script execution timeout (60s)")
        except FileNotFoundError as e:
            if "python" in str(e).lower():
                raise RuntimeError("Python no está instalado o no está en el PATH")
            raise
    
    def download(self, xlsx_path: str, custom_name: str = None) -> str:
        """Renombra el archivo Excel generado con un nombre descriptivo.
        
        Args:
            xlsx_path: Ruta del archivo generado
            custom_name: Nombre personalizado (opcional). Si no se proporciona,
                        se genera uno basado en el contenido usando IA
            
        Returns:
            Nueva ruta del archivo renombrado
        """
        xlsx_file = Path(xlsx_path)
        
        if not xlsx_file.exists():
            raise FileNotFoundError(f"Archivo no encontrado: {xlsx_path}")
        
        # Si no se proporciona nombre, generar uno usando IA
        if not custom_name:
            custom_name = self._generate_filename(xlsx_path)
        
        # Limpiar el nombre (remover caracteres no válidos)
        custom_name = self._sanitize_filename(custom_name)
        
        # Asegurar extensión .xlsx
        if not custom_name.endswith('.xlsx'):
            custom_name += '.xlsx'
        
        # Nueva ruta
        new_path = self.output_dir / custom_name
        
        # Renombrar
        xlsx_file.rename(new_path)
        
        print(f"[ExcelScriptGenerator] Archivo renombrado: {new_path}")
        return str(new_path)
    
    def _generate_filename(self, xlsx_path: str) -> str:
        """Genera un nombre descriptivo para el archivo usando IA."""
        prompt = (
            "Basándote en el contexto de la conversación anterior sobre el archivo Excel, "
            "genera un nombre de archivo corto y descriptivo (máximo 50 caracteres, sin espacios, usa guiones). "
            "Responde SOLO con el nombre del archivo, sin extensión ni explicaciones."
        )
        
        filename = self.engine.process(prompt).strip()
        
        # Limpiar respuesta
        filename = filename.replace('"', '').replace("'", '').strip()
        
        return filename
    
    def _sanitize_filename(self, filename: str) -> str:
        """Limpia el nombre de archivo removiendo caracteres no válidos."""
        # Remover extensión si existe
        if filename.endswith('.xlsx'):
            filename = filename[:-5]
        
        # Reemplazar caracteres no válidos
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '-')
        
        # Reemplazar espacios con guiones
        filename = filename.replace(' ', '-')
        
        # Limitar longitud
        if len(filename) > 50:
            filename = filename[:50]
        
        return filename
    
    def generate_and_execute(self, user_request: str) -> tuple[str, str]:
        """Genera y ejecuta el script en un solo paso.
        
        Args:
            user_request: Descripción del archivo Excel que el usuario quiere generar
            
        Returns:
            Tupla (script_path, xlsx_path)
        """
        script_path = self.generate_script(user_request)
        xlsx_path = self.execute_script(script_path)
        return script_path, xlsx_path
