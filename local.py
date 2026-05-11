import os
import sys
import flet as ft
from colorama import init, Fore, Style

init(autoreset=True)

# Agregar directorios al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "assets"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "doc"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "xlsx"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "pptx"))

from assets.engine import OpenRouterEngine
from assets.openrouter import OpenRouterCatalog
from word import WordScriptGenerator
from excel import ExcelScriptGenerator
from powerpoint import PowerPointScriptGenerator


def main():
    """Interfaz gráfica con Flet para chatear con modelos de IA y generar documentos."""
    
    def app_main(page: ft.Page):
        page.title = "GR Docs - Chat IA & Generador"
        page.theme_mode = ft.ThemeMode.DARK
        page.padding = 20
        page.window_width = 1100
        page.window_height = 750
        
        # Estado
        catalog = OpenRouterCatalog()
        free_models = catalog.filter_models(free_only=True, modality="text->text")
        current_engine = None
        
        # ─── COMPONENTES ──────────────────────────────────────────────────────
        
        # Dropdown de modelos
        model_dropdown = ft.Dropdown(
            label="Modelo de IA",
            hint_text="Elige un modelo",
            options=[
                ft.dropdown.Option(
                    key=m["id"],
                    text=f"{m['name'][:40]}..." if len(m['name']) > 40 else m['name']
                )
                for m in free_models
            ],
            width=400,
            value=free_models[0]["id"] if free_models else None
        )
        
        # Área de chat
        chat_container = ft.ListView(
            expand=True,
            spacing=10,
            padding=20,
            auto_scroll=True
        )
        
        # Input de mensaje
        message_input = ft.TextField(
            hint_text="Escribe tu mensaje o describe el documento que quieres generar...",
            multiline=True,
            min_lines=2,
            max_lines=6,
            expand=True,
            autofocus=True,
            on_submit=lambda e: send_message()
        )
        
        # Botones
        send_button = ft.ElevatedButton(
            "Chat",
            icon=ft.icons.CHAT,
            on_click=lambda e: send_message(),
            bgcolor=ft.colors.BLUE_700
        )
        
        generate_word_btn = ft.ElevatedButton(
            "Word",
            icon=ft.icons.DESCRIPTION,
            bgcolor=ft.colors.INDIGO_700,
            on_click=lambda e: generate_document("word"),
            tooltip="Generar documento Word"
        )
        
        generate_excel_btn = ft.ElevatedButton(
            "Excel",
            icon=ft.icons.TABLE_CHART,
            bgcolor=ft.colors.GREEN_700,
            on_click=lambda e: generate_document("excel"),
            tooltip="Generar archivo Excel"
        )
        
        generate_pptx_btn = ft.ElevatedButton(
            "PowerPoint",
            icon=ft.icons.SLIDESHOW,
            bgcolor=ft.colors.ORANGE_700,
            on_click=lambda e: generate_document("pptx"),
            tooltip="Generar presentación PowerPoint"
        )
        
        clear_button = ft.IconButton(
            icon=ft.icons.DELETE_SWEEP,
            tooltip="Limpiar chat",
            on_click=lambda e: clear_chat()
        )
        
        loading_indicator = ft.ProgressRing(visible=False, width=20, height=20)
        
        # ─── FUNCIONES ────────────────────────────────────────────────────────
        
        def initialize_engine():
            """Inicializa el engine con el modelo seleccionado."""
            nonlocal current_engine
            
            selected_model = model_dropdown.value
            if not selected_model:
                return False
            
            try:
                import yaml
                with open("settings.yaml") as f:
                    settings = yaml.safe_load(f)
                
                original_model = settings.get("model")
                settings["model"] = selected_model
                
                with open("settings.yaml", "w") as f:
                    yaml.dump(settings, f)
                
                current_engine = OpenRouterEngine(verbose=False)
                
                settings["model"] = original_model
                with open("settings.yaml", "w") as f:
                    yaml.dump(settings, f)
                
                return True
            except Exception as e:
                show_error(f"Error al inicializar modelo: {e}")
                return False
        
        def add_message(content: str, is_user: bool, message_type: str = "text"):
            """Agrega un mensaje al chat."""
            
            if message_type == "info":
                icon = ft.icons.INFO
                color = ft.colors.BLUE_400
                bg_color = ft.colors.BLUE_900
            elif message_type == "success":
                icon = ft.icons.CHECK_CIRCLE
                color = ft.colors.GREEN_400
                bg_color = ft.colors.GREEN_900
            elif message_type == "error":
                icon = ft.icons.ERROR
                color = ft.colors.RED_400
                bg_color = ft.colors.RED_900
            else:
                icon = ft.icons.PERSON if is_user else ft.icons.SMART_TOY
                color = ft.colors.BLUE_400 if is_user else ft.colors.GREEN_400
                bg_color = ft.colors.BLUE_GREY_900 if is_user else ft.colors.BLUE_GREY_800
            
            message_card = ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(icon, color=color),
                        ft.Text(
                            "Tú" if is_user and message_type == "text" else 
                            "IA" if not is_user and message_type == "text" else
                            "Sistema",
                            weight=ft.FontWeight.BOLD,
                            color=color
                        )
                    ]),
                    ft.Text(content, selectable=True)
                ]),
                bgcolor=bg_color,
                border_radius=10,
                padding=15,
                margin=ft.margin.only(
                    left=50 if is_user and message_type == "text" else 0,
                    right=0 if is_user and message_type == "text" else 50
                )
            )
            
            chat_container.controls.append(message_card)
            page.update()
        
        def show_error(message: str):
            """Muestra un mensaje de error."""
            add_message(message, False, "error")
        
        def show_info(message: str):
            """Muestra un mensaje informativo."""
            add_message(message, False, "info")
        
        def show_success(message: str):
            """Muestra un mensaje de éxito."""
            add_message(message, False, "success")
        
        def send_message():
            """Envía un mensaje al modelo de IA."""
            user_message = message_input.value.strip()
            
            if not user_message:
                return
            
            message_input.value = ""
            message_input.focus()
            
            add_message(user_message, is_user=True)
            
            loading_indicator.visible = True
            send_button.disabled = True
            page.update()
            
            try:
                if current_engine is None:
                    if not initialize_engine():
                        loading_indicator.visible = False
                        send_button.disabled = False
                        page.update()
                        return
                
                response = current_engine.process(user_message)
                add_message(response, is_user=False)
                
            except Exception as e:
                show_error(f"Error: {str(e)}")
            
            finally:
                loading_indicator.visible = False
                send_button.disabled = False
                page.update()
        
        def generate_document(doc_type: str):
            """Genera un documento Word, Excel o PowerPoint."""
            user_request = message_input.value.strip()
            
            if not user_request:
                show_error("Por favor describe el documento que quieres generar")
                return
            
            message_input.value = ""
            
            # Mostrar solicitud
            add_message(f"📄 Generando {doc_type.upper()}: {user_request}", True)
            
            # Deshabilitar botones
            loading_indicator.visible = True
            generate_word_btn.disabled = True
            generate_excel_btn.disabled = True
            generate_pptx_btn.disabled = True
            page.update()
            
            try:
                if doc_type == "word":
                    show_info("Generando documento Word...")
                    generator = WordScriptGenerator(verbose=False)
                    script_path, doc_path = generator.generate_and_execute(user_request)
                    final_path = generator.download(doc_path)
                    show_success(f"✓ Documento Word generado:\n{final_path}")
                
                elif doc_type == "excel":
                    show_info("Generando archivo Excel...")
                    generator = ExcelScriptGenerator(verbose=False)
                    script_path, doc_path = generator.generate_and_execute(user_request)
                    final_path = generator.download(doc_path)
                    show_success(f"✓ Archivo Excel generado:\n{final_path}")
                
                elif doc_type == "pptx":
                    show_info("Generando presentación PowerPoint...")
                    generator = PowerPointScriptGenerator(verbose=False)
                    script_path, doc_path = generator.generate_and_execute(user_request)
                    final_path = generator.download(doc_path)
                    show_success(f"✓ Presentación PowerPoint generada:\n{final_path}")
                
            except Exception as e:
                show_error(f"Error al generar {doc_type}: {str(e)}")
            
            finally:
                loading_indicator.visible = False
                generate_word_btn.disabled = False
                generate_excel_btn.disabled = False
                generate_pptx_btn.disabled = False
                page.update()
        
        def clear_chat():
            """Limpia el historial del chat."""
            nonlocal current_engine
            chat_container.controls.clear()
            if current_engine:
                current_engine.clear_history()
            
            # Mensaje de bienvenida
            add_welcome_message()
            page.update()
        
        def change_model(e):
            """Cambia el modelo de IA."""
            nonlocal current_engine
            current_engine = None
            
            selected_name = next(
                (m["name"] for m in free_models if m["id"] == model_dropdown.value),
                "Desconocido"
            )
            show_info(f"Modelo cambiado a: {selected_name}")
        
        def add_welcome_message():
            """Agrega mensaje de bienvenida."""
            welcome_card = ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.icons.WAVING_HAND_ROUNDED, color=ft.colors.AMBER_400, size=30),
                        ft.Text(
                            "¡Bienvenido a GR Docs!",
                            size=20,
                            weight=ft.FontWeight.BOLD,
                            color=ft.colors.AMBER_400
                        )
                    ]),
                    ft.Divider(height=10, color=ft.colors.TRANSPARENT),
                    ft.Text(
                        f"🤖 {len(free_models)} modelos gratuitos disponibles",
                        color=ft.colors.WHITE70
                    ),
                    ft.Text(
                        "💬 Usa el botón 'Chat' para conversar con la IA",
                        color=ft.colors.WHITE70
                    ),
                    ft.Text(
                        "📄 Usa los botones Word/Excel/PowerPoint para generar documentos",
                        color=ft.colors.WHITE70
                    ),
                    ft.Divider(height=10, color=ft.colors.TRANSPARENT),
                    ft.Text(
                        "Ejemplo: 'Crea un informe de ventas Q1 2024 con gráficos' → Word",
                        size=12,
                        italic=True,
                        color=ft.colors.WHITE54
                    )
                ]),
                bgcolor=ft.colors.AMBER_900,
                border_radius=10,
                padding=20
            )
            chat_container.controls.append(welcome_card)
        
        model_dropdown.on_change = change_model
        
        # ─── LAYOUT ───────────────────────────────────────────────────────────
        
        page.add(
            ft.Column([
                # Header
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.icons.DESCRIPTION, size=40, color=ft.colors.BLUE_400),
                        ft.Text(
                            "GR Docs - Chat & Generador",
                            size=24,
                            weight=ft.FontWeight.BOLD
                        ),
                        ft.Container(expand=True),
                        model_dropdown,
                        clear_button
                    ]),
                    padding=10,
                    bgcolor=ft.colors.BLUE_GREY_900,
                    border_radius=10
                ),
                
                # Chat area
                ft.Container(
                    content=chat_container,
                    expand=True,
                    bgcolor=ft.colors.BLUE_GREY_900,
                    border_radius=10,
                    padding=10
                ),
                
                # Input area
                ft.Container(
                    content=ft.Column([
                        message_input,
                        ft.Row([
                            loading_indicator,
                            send_button,
                            ft.VerticalDivider(width=20, color=ft.colors.TRANSPARENT),
                            ft.Text("Generar:", color=ft.colors.WHITE70),
                            generate_word_btn,
                            generate_excel_btn,
                            generate_pptx_btn
                        ])
                    ]),
                    padding=10,
                    bgcolor=ft.colors.BLUE_GREY_900,
                    border_radius=10
                )
            ], expand=True)
        )
        
        # Mensaje de bienvenida inicial
        add_welcome_message()
        page.update()
    
    # Iniciar app
    ft.app(target=app_main)


if __name__ == "__main__":
    main()
