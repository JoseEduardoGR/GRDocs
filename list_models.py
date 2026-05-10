#!/usr/bin/env python3
"""Script para listar todos los modelos gratuitos disponibles en OpenRouter."""

import sys
import os
from colorama import init, Fore, Style

# Agregar el directorio assets al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "assets"))

from openrouter import OpenRouterCatalog

init(autoreset=True)


def main():
    print(Fore.CYAN + Style.BRIGHT + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + Style.BRIGHT + "║     MODELOS GRATUITOS DISPONIBLES EN OPENROUTER               ║")
    print(Fore.CYAN + Style.BRIGHT + "╚════════════════════════════════════════════════════════════════╝\n")
    
    catalog = OpenRouterCatalog()
    
    # Obtener modelos de texto
    modelos_texto = catalog.filter_models(free_only=True, modality="text->text")
    
    # Obtener modelos con visión
    modelos_vision = catalog.filter_models(free_only=True, modality="text+image->text")
    
    # ── MODELOS DE TEXTO ──────────────────────────────────────────────────────
    print(Fore.YELLOW + Style.BRIGHT + f"📝 MODELOS DE TEXTO ({len(modelos_texto)} disponibles)")
    print(Fore.YELLOW + "─" * 80 + "\n")
    
    for i, m in enumerate(modelos_texto, 1):
        print(Fore.GREEN + f"{i}. {m['name']}")
        print(Fore.WHITE + f"   ID: {m['id']}")
        print(Fore.CYAN + f"   Contexto: {m['context']:,} tokens")
        
        # Descripción (truncada si es muy larga)
        desc = m['description']
        if len(desc) > 150:
            desc = desc[:150] + "..."
        print(Fore.WHITE + f"   Descripción: {desc}")
        print(Fore.BLUE + f"   🔗 {m['link']}")
        print()
    
    # ── MODELOS CON VISIÓN ────────────────────────────────────────────────────
    print(Fore.YELLOW + Style.BRIGHT + f"\n👁️  MODELOS CON VISIÓN ({len(modelos_vision)} disponibles)")
    print(Fore.YELLOW + "─" * 80 + "\n")
    
    for i, m in enumerate(modelos_vision, 1):
        print(Fore.GREEN + f"{i}. {m['name']}")
        print(Fore.WHITE + f"   ID: {m['id']}")
        print(Fore.CYAN + f"   Contexto: {m['context']:,} tokens")
        
        # Descripción (truncada si es muy larga)
        desc = m['description']
        if len(desc) > 150:
            desc = desc[:150] + "..."
        print(Fore.WHITE + f"   Descripción: {desc}")
        print(Fore.BLUE + f"   🔗 {m['link']}")
        print()
    
    # ── RESUMEN ───────────────────────────────────────────────────────────────
    print(Fore.CYAN + Style.BRIGHT + "\n╔════════════════════════════════════════════════════════════════╗")
    print(Fore.CYAN + Style.BRIGHT + f"║  TOTAL: {len(modelos_texto)} modelos de texto + {len(modelos_vision)} modelos con visión")
    print(Fore.CYAN + Style.BRIGHT + "╚════════════════════════════════════════════════════════════════╝\n")
    
    print(Fore.YELLOW + "💡 Tip: Copia el ID del modelo que quieras usar en tu código")
    print(Fore.WHITE + "   Ejemplo: engine.model = 'google/gemini-2.0-flash-001:free'\n")


if __name__ == "__main__":
    main()
