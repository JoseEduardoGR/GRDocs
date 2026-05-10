# para activar el entorno es "source venv/bin/activate.fish"

import argparse
import yaml
import local
import server
from colorama import init, Fore, Style

# Handles ANSI color codes on Windows automatically, no-op on Linux/macOS
init(autoreset=True)

with open("settings.yaml") as f:
    settings = yaml.safe_load(f)

parser = argparse.ArgumentParser()
parser.add_argument('-m', '--mode', choices=['local', 'l', 'api', 'a'], default=None,
                    help='Choose "local/l" for the GUI or "api/a" for server-side integration.')
args = parser.parse_args()

MODES = {
    ('local', 'l'): local.main,
    ('api',   'a'): server.main,
}

def get_handler(mode):
    for keys, handler in MODES.items():
        if mode in keys:
            return handler
    return None

def show_welcome():
    print(Fore.CYAN + '''
#=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=#

░██████╗░██████╗░░░░░██████╗░░█████╗░░█████╗░░██████╗
██╔════╝░██╔══██╗░░░░██╔══██╗██╔══██╗██╔══██╗██╔════╝
██║░░██╗░██████╔╝░░░░██║░░██║██║░░██║██║░░╚═╝╚█████╗░
██║░░╚██╗██╔══██╗░░░░██║░░██║██║░░██║██║░░██╗░╚═══██╗
╚██████╔╝██║░░██║░░░░██████╔╝╚█████╔╝╚█████╔╝██████╔╝
░╚═════╝░╚═╝░░╚═╝░░░░╚═════╝░░╚════╝░░╚════╝░╚═════╝░

#=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=#
    ''')
    print(Fore.YELLOW + "Welcome! Please select a mode to get started:")
    print(Fore.GREEN  + "  -m local  or  --mode local" + Style.RESET_ALL + "   Launch the graphical user interface")
    print(Fore.GREEN  + "  -m l"                       + Style.RESET_ALL + "                         Short version for local mode")
    print(Fore.BLUE   + "  -m api   or  --mode api"    + Style.RESET_ALL + "      Start the server-side API integration")

# Resolve mode: CLI arg takes priority, then settings.yaml, then show welcome
mode = args.mode or settings.get('default_mode')
handler = get_handler(mode) if mode else None

if handler:
    handler()
else:
    show_welcome()
