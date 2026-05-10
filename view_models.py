import os
import sys
import threading

import flet as ft

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "assets"))

from engine import OpenRouterEngine
from openrouter import OpenRouterCatalog

# ─── Tokenización simple (aprox. 1 token ≈ 4 chars) ──────────────────────────

def estimate_tokens(text: str) -> int:
    return max(1, len(text) // 4)


def history_tokens(history: list) -> int:
    total = 0
    for msg in history:
        content = msg.get("content", "")
        if isinstance(content, list):
            for part in content:
                if part.get("type") == "text":
                    total += estimate_tokens(part["text"])
        else:
            total += estimate_tokens(str(content))
    return total


# ─── Colores ──────────────────────────────────────────────────────────────────

BG       = "#0f1117"
SURFACE  = "#1a1d27"
SURFACE2 = "#22263a"
ACCENT   = "#7c6af7"
ACCENT2  = "#a78bfa"
TEXT     = "#e2e8f0"
SUBTEXT  = "#94a3b8"
USER_BG  = "#2d2250"
BOT_BG   = "#1e2235"
DANGER   = "#f87171"
SUCCESS  = "#34d399"
WARNING  = "#fbbf24"


# ─── App ──────────────────────────────────────────────────────────────────────

def main(page: ft.Page):
    page.title      = "GR Docs · Models"
    page.bgcolor    = BG
    page.window.width  = 1000
    page.window.height = 720
    page.padding    = 0
    page.theme_mode = ft.ThemeMode.DARK

    # ── Estado ────────────────────────────────────────────────────────────────
    engine: OpenRouterEngine | None = None
    all_models: list[dict]          = []
    selected_model: dict | None     = None
    pending_image: str | None       = None

    # ── Refs selector ─────────────────────────────────────────────────────────
    catalog_loading = ft.Ref[ft.ProgressRing]()
    model_dropdown  = ft.Ref[ft.DropdownM2]()
    model_info_text = ft.Ref[ft.Text]()

    # ── Refs chat ─────────────────────────────────────────────────────────────
    chat_column  = ft.Ref[ft.Column]()
    input_field  = ft.Ref[ft.TextField]()
    send_btn     = ft.Ref[ft.IconButton]()
    attach_btn   = ft.Ref[ft.IconButton]()
    token_bar    = ft.Ref[ft.ProgressBar]()
    token_label  = ft.Ref[ft.Text]()
    image_chip   = ft.Ref[ft.Container]()
    thinking_row = ft.Ref[ft.Row]()

    # ── Vista principal ───────────────────────────────────────────────────────
    view_selector = ft.Ref[ft.AnimatedSwitcher]()

    # ══════════════════════════════════════════════════════════════════════════
    #  HELPERS
    # ══════════════════════════════════════════════════════════════════════════

    def update_token_bar(extra_text: str = ""):
        if engine is None or selected_model is None:
            return
        ctx_limit = selected_model.get("context", 4096) or 4096
        used      = history_tokens(engine.history) + estimate_tokens(extra_text)
        ratio     = min(used / ctx_limit, 1.0)
        remaining = max(ctx_limit - used, 0)

        token_bar.current.value = ratio
        token_bar.current.color = (
            SUCCESS if ratio < 0.6 else
            WARNING if ratio < 0.85 else
            DANGER
        )
        token_label.current.value = (
            f"{remaining:,} tokens restantes  ({used:,} / {ctx_limit:,})"
        )
        page.update()

    def add_message(role: str, text: str, img_path: str = None):
        is_user  = role == "user"
        bg_color = USER_BG if is_user else BOT_BG
        align    = ft.CrossAxisAlignment.END if is_user else ft.CrossAxisAlignment.START

        parts: list[ft.Control] = []

        if img_path and os.path.exists(img_path):
            parts.append(ft.Image(src=img_path, width=220, border_radius=8, fit=ft.BoxFit.CONTAIN))

        parts.append(ft.Text(value=text, color=TEXT, size=14, selectable=True))

        bubble = ft.Container(
            content=ft.Column(parts, spacing=6, tight=True),
            bgcolor=bg_color,
            border_radius=12,
            padding=ft.Padding(left=14, right=14, top=10, bottom=10),
            margin=ft.Margin(
                left=60 if is_user else 0,
                right=0  if is_user else 60,
                top=0,
                bottom=0,
            ),
        )

        label = ft.Text(
            value="Tú" if is_user else "Modelo",
            color=ACCENT2 if is_user else SUBTEXT,
            size=11,
            weight=ft.FontWeight.W_600,
        )

        row = ft.Column(
            controls=[label, bubble],
            horizontal_alignment=align,
            spacing=3,
        )

        chat_column.current.controls.append(row)
        page.update()

    def set_thinking(visible: bool):
        thinking_row.current.visible = visible
        send_btn.current.disabled    = visible
        input_field.current.disabled = visible
        attach_btn.current.disabled  = visible
        page.update()

    def clear_image_chip():
        nonlocal pending_image
        pending_image = None
        image_chip.current.visible = False
        page.update()

    # ══════════════════════════════════════════════════════════════════════════
    #  EVENTOS
    # ══════════════════════════════════════════════════════════════════════════

    def on_input_change(e):
        update_token_bar(extra_text=e.control.value)

    def on_send(e):
        nonlocal pending_image
        text = input_field.current.value.strip()
        if not text or engine is None:
            return

        img = pending_image
        input_field.current.value = ""
        clear_image_chip()
        set_thinking(True)
        add_message("user", text, img_path=img)

        def run():
            response = engine.process(text, image_path=img)
            add_message("assistant", response)
            set_thinking(False)
            update_token_bar()

        threading.Thread(target=run, daemon=True).start()

    def on_attach_result(e: ft.FilePickerResultEvent):
        nonlocal pending_image
        if not e.files:
            return
        pending_image = e.files[0].path
        fname = os.path.basename(pending_image)

        close_btn = ft.IconButton(icon=ft.Icons.CLOSE, icon_size=14, icon_color=SUBTEXT, tooltip="Quitar imagen")
        close_btn.on_click   = lambda _: clear_image_chip()

        chip_row = ft.Row(spacing=4, tight=True)
        chip_row.controls = [
            ft.Icon(ft.Icons.IMAGE, color=ACCENT2, size=16),
            ft.Text(fname, color=ACCENT2, size=12,
                    max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
            close_btn,
        ]
        image_chip.current.content = chip_row
        image_chip.current.visible = True
        page.update()

    # FilePicker: en flet 0.85 va en page.services, no en page.overlay
    file_picker = ft.FilePicker()
    file_picker.on_result = on_attach_result
    page.services.append(file_picker)
    page.update()

    def on_attach(e):
        file_picker.pick_files(
            dialog_title="Selecciona una imagen",
            allowed_extensions=["png", "jpg", "jpeg", "gif", "webp"],
            allow_multiple=False,
        )

    def on_clear_chat(e):
        if engine:
            engine.clear_history()
        chat_column.current.controls.clear()
        update_token_bar()
        page.update()

    def on_model_change(e):
        nonlocal selected_model
        val = e.control.value
        selected_model = next((m for m in all_models if m["id"] == val), None)
        if selected_model:
            on_start_chat(None)
        else:
            model_info_text.current.value = ""
            page.update()

    def on_start_chat(e):
        nonlocal engine
        if selected_model is None:
            return
        engine              = OpenRouterEngine()
        engine.model        = selected_model["id"]
        engine.model_vision = selected_model["id"]
        view_selector.current.content = build_chat_view()
        update_token_bar()
        page.update()

    def on_back_to_selector(e):
        nonlocal engine
        engine = None
        view_selector.current.content = build_selector_view()
        page.update()

    # ══════════════════════════════════════════════════════════════════════════
    #  VISTAS
    # ══════════════════════════════════════════════════════════════════════════

    def build_selector_view() -> ft.Control:
        dropdown = ft.DropdownM2(
            ref=model_dropdown,
            hint_text="Cargando modelos...",
            disabled=True,
            on_change=on_model_change,
            bgcolor=SURFACE2,
            border_color=ACCENT,
            color=TEXT,
            focused_border_color=ACCENT2,
            expand=True,
        )

        ring = ft.ProgressRing(
            ref=catalog_loading,
            width=20, height=20,
            stroke_width=2,
            color=ACCENT2,
        )

        return ft.Container(
            expand=True,
            bgcolor=BG,
            content=ft.Column(
                expand=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=24,
                controls=[
                    ft.Column(
                        [
                            ft.Text("GR Docs", size=36,
                                    weight=ft.FontWeight.BOLD, color=ACCENT2),
                            ft.Text("Selecciona un modelo para comenzar",
                                    size=14, color=SUBTEXT),
                        ],
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=4,
                    ),
                    ft.Container(
                        width=520,
                        bgcolor=SURFACE,
                        border_radius=16,
                        padding=ft.Padding(28),
                        content=ft.Column(
                            spacing=16,
                            controls=[
                                ft.Text("Modelo", size=13, color=SUBTEXT,
                                        weight=ft.FontWeight.W_500),
                                ft.Stack(controls=[
                                    dropdown,
                                    ft.Container(content=ring, right=10, top=10),
                                ]),
                                ft.Text(ref=model_info_text, value="",
                                        size=12, color=SUBTEXT),
                            ],
                        ),
                    ),
                ],
            ),
        )

    def build_chat_view() -> ft.Control:
        name = selected_model["name"] if selected_model else "Modelo"
        ctx  = selected_model.get("context", 4096) if selected_model else 4096

        # IconButtons: asignar eventos como atributos
        back_btn = ft.IconButton(icon=ft.Icons.ARROW_BACK_IOS_NEW, icon_color=SUBTEXT, tooltip="Cambiar modelo")
        back_btn.on_click = on_back_to_selector

        clear_btn = ft.IconButton(icon=ft.Icons.DELETE_OUTLINE, icon_color=SUBTEXT, tooltip="Limpiar chat")
        clear_btn.on_click = on_clear_chat

        attach_ib = ft.IconButton(
            ref=attach_btn, icon=ft.Icons.ATTACH_FILE,
            icon_color=SUBTEXT, tooltip="Adjuntar imagen"
        )
        attach_ib.on_click = on_attach

        send_ib = ft.IconButton(
            ref=send_btn, icon=ft.Icons.SEND_ROUNDED,
            icon_color=ACCENT2, tooltip="Enviar (Enter)"
        )
        send_ib.on_click = on_send

        tf = ft.TextField(
            ref=input_field,
            hint_text="Escribe un mensaje...",
            hint_style=ft.TextStyle(color=SUBTEXT),
            bgcolor=SURFACE2,
            border_color="transparent",
            focused_border_color=ACCENT,
            color=TEXT,
            cursor_color=ACCENT2,
            multiline=True,
            min_lines=1,
            max_lines=5,
            expand=True,
            shift_enter=True,
        )
        tf.on_change = on_input_change
        tf.on_submit = on_send

        return ft.Column(
            expand=True,
            spacing=0,
            controls=[
                # Header
                ft.Container(
                    bgcolor=SURFACE,
                    padding=ft.Padding(left=20, right=20, top=12, bottom=12),
                    content=ft.Row(
                        [
                            back_btn,
                            ft.Column(
                                [
                                    ft.Text(name, size=15,
                                            weight=ft.FontWeight.W_600, color=TEXT),
                                    ft.Text(f"Contexto: {ctx:,} tokens",
                                            size=11, color=SUBTEXT),
                                ],
                                spacing=1,
                                expand=True,
                            ),
                            clear_btn,
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                ),

                # Barra de tokens
                ft.Container(
                    bgcolor=SURFACE2,
                    padding=ft.Padding(left=20, right=20, top=6, bottom=6),
                    content=ft.Column(
                        [
                            ft.ProgressBar(
                                ref=token_bar,
                                value=0,
                                color=SUCCESS,
                                bgcolor="#2a2f45",
                                height=6,
                                border_radius=3,
                            ),
                            ft.Text(
                                ref=token_label,
                                value=f"{ctx:,} tokens restantes  (0 / {ctx:,})",
                                size=11,
                                color=SUBTEXT,
                            ),
                        ],
                        spacing=4,
                    ),
                ),

                # Mensajes
                ft.Container(
                    expand=True,
                    bgcolor=BG,
                    padding=ft.Padding(left=20, right=20, top=14, bottom=14),
                    content=ft.Column(
                        ref=chat_column,
                        expand=True,
                        scroll=ft.ScrollMode.AUTO,
                        spacing=10,
                        controls=[],
                        auto_scroll=True,
                    ),
                ),

                # Indicador "pensando"
                ft.Row(
                    ref=thinking_row,
                    visible=False,
                    spacing=0,
                    controls=[
                        ft.Container(width=20),
                        ft.ProgressRing(width=16, height=16,
                                        stroke_width=2, color=ACCENT2),
                        ft.Text("  Generando respuesta...", size=12, color=SUBTEXT),
                    ],
                ),

                # Chip imagen adjunta
                ft.Container(
                    ref=image_chip,
                    visible=False,
                    bgcolor=SURFACE2,
                    border_radius=8,
                    padding=ft.Padding(left=12, right=12, top=4, bottom=4),
                    margin=ft.Margin(left=16, right=16, top=0, bottom=4),
                ),

                # Input
                ft.Container(
                    bgcolor=SURFACE,
                    padding=ft.Padding(left=12, right=12, top=10, bottom=10),
                    content=ft.Row(
                        [attach_ib, tf, send_ib],
                        vertical_alignment=ft.CrossAxisAlignment.END,
                        spacing=4,
                    ),
                ),
            ],
        )

    # ══════════════════════════════════════════════════════════════════════════
    #  LAYOUT PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════

    page.add(
        ft.AnimatedSwitcher(
            ref=view_selector,
            content=build_selector_view(),
            transition=ft.AnimatedSwitcherTransition.FADE,
            duration=200,
            expand=True,
        )
    )

    # Carga de modelos en background
    def load_models():
        nonlocal all_models
        catalog       = OpenRouterCatalog()
        text_models   = catalog.filter_models(free_only=True, modality="text->text")
        vision_models = catalog.filter_models(free_only=True, modality="text+image->text")

        seen, combined = set(), []
        for m in vision_models:
            if m["id"] not in seen:
                m["_vision"] = True
                combined.append(m)
                seen.add(m["id"])
        for m in text_models:
            if m["id"] not in seen:
                m["_vision"] = False
                combined.append(m)
                seen.add(m["id"])

        all_models = combined

        options = [
            ft.dropdownm2.Option(
                key=m["id"],
                text=f"{'👁 ' if m['_vision'] else '💬 '}{m['name']}  ({m['context']:,} ctx)",
            )
            for m in all_models
        ]

        model_dropdown.current.options   = options
        model_dropdown.current.hint_text = "Elige un modelo..."
        model_dropdown.current.disabled  = False
        catalog_loading.current.visible  = False
        page.update()

    threading.Thread(target=load_models, daemon=True).start()


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ft.run(main)