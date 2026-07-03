# ======================================================
# PEAT
# just peat
# Made by Beffy
# ======================================================

# TODO list (arrows mean WIP):
# - Add extension JSON manifest
# - Add different 'modules', such as libraries, apis, and include extensions
# - Add an Action Log for P.E.A.T. and user to keep track of what it does, force extensions to log their actions
# - Differentiate between 'System Log' and 'Action Log'

# TODO PBAT:
# - Fix PBAT loops
# - Add PBAT recursion (PBAT in PBAT)
# - Add trusted PBAT folder system, or json
# - Add warning prompt before running untrusted PBAT scripts
# - Add safe mode for dangerous commands
# - Add optional command permission system
# - Add PBAT execution restrictions/sandboxing

# TODO later list:
# - Add more error codes and better error handling
# - Add executables (not windows executables, custom shortscript maybe)

# TODO ideas:
# - Add turtle faces
# - Add diagnostics

from datetime import datetime
from vosk import Model, KaldiRecognizer

import sounddevice as sd
import numpy as np
import win32com.client
import json
import queue
import cv2

import importlib.util
import time
import sys
import os
import shutil
import subprocess
import webbrowser
import shlex
import difflib


# ====== Variables ======
# == Constants ==
PEAT_VERSION = ""
MAX_REPEAT = 60
MAX_LABEL_LENGTH = 24
PBAT_MAX_REPEAT = 60
PBAT_MAX_LABEL_LEN = 24
PBAT_RESERVED = {"end"}
SAMPLE_RATE = 16000
SILENCE_LIMIT = 0
BOT_NAME = "Peat"
BOT_NAME_MADE_WITH_AI = False

# == Init ==
# Nothing here... for now

# == Defaults ==
default_name = "User"
fallback_input = "wubbaLubbaDubDubs" # Yup
default_query_placeholder = "How can I assist you? "

tts_toggle = True
debug_info = False

q = queue.Queue()

voice_model = None

# == User ==
name = default_name

# == TTS ==
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 1
speaker.Volume = 100
samplerate = 16000

# == Command Variables ==
camera_mode = False
camera_index = 0

# == Status ==
status = 0
error_code = 0
voice_ready = True

# == Stats ==
total_commands_used = 0

# == Developer ==
debug_mode = False
dev_mode = False

# == Paths ==
HOME = os.path.expanduser("~")
BASE_DIR = os.path.dirname(os.path.realpath(__file__))
EXT_ROOT = os.path.join(BASE_DIR, "PEAT_Modules/Extensions")
EXT_STOCK = os.path.join(EXT_ROOT, "Stock")
EXT_UNACTIVE = os.path.join(EXT_ROOT, "Inactive")
EXT_ACTIVE = os.path.join(EXT_ROOT, "Active")

config_folder = os.path.join(BASE_DIR, "config")
config_path = os.path.join(config_folder, "config.json")
pbat_folder = os.path.join(BASE_DIR, "PEAT_Batch")
log_dump_folder = os.path.join(BASE_DIR, "log_dumps")
camera_print_path = os.path.join(HOME, "Pictures", "PEAT Cam Captures")
user_info_path = os.path.join(config_folder, "user_info.json")
pbat_trust_folder = os.path.join(BASE_DIR, "PEAT_Batch/Trusted")
system_folder = os.path.join(BASE_DIR, "system")
sys_info_path = os.path.join(system_folder, "sys_info.json")
peat_print_path = os.path.join(HOME, "Documents", "PEAT Prints")

# == Files ==
log_file = "peat_log.txt"

# == Everything Else ==
query = ""
query_placeholder = default_query_placeholder

# === Lists ===
executables = [
    "temp",
    "temp2"
]

quotes = [
    "Nobody exists on purpose. Nobody belongs anywhere. Everybody's gonna die. Come watch TV",
    "dude stop showing your chest bruh [out of context]",
    "Welcome to the club, pal.",
    "Cow urine cures covid",
    "Everyone else is stupid and I'm the only one smart"
]

# == Dictionaries ==
router = {}
extensions = {}
cam_controls = {}
user_info = {}
extension_help = {}
loaded_extensions = []

help_pages = {
    "help": "Displays this help message or details for a specific command.",
    "launch": "Launches the specified application name.",
    "exec": "Does nothing, is a work-in-progress.",
    "pbat": "Runs a PEAT Batch (.pbat) script from the 'PEAT Batch' folder.",
    "time": "Displays the current local time.",
    "date": "Displays the current local date.",
    "now": "Displays the current local date and time.",
    "log": "Inspects or modifies debug logs.",
    "config.get": "[ADVANCED] Reads a value from the config file.",
    "config.set": "[ADVANCED] Writes a value to the config file.",
    "delay": "Waits the specified number of seconds before continuing.",
    "link": "Opens the specified URL in the default web browser.",
    "print": "Prints the provided string.",
    "quote": "Displays a random quote from the quote list.",
    "quit": "Exits PEAT.",
}

voice_aliases = {
    "peat": ["peat", "pete", "beat", "beet", "petey", "pet"],
    "help": ["help", "halp", "held", "health", "help me", "help me please"],
    "ask": ["ask", "question", "query", "ask question"],
    "launch": ["launch", "open", "start", "run app", "start app"],
    "exec": ["exec", "execute", "run", "run command", "execute command"],
    "pbat": ["pbat", "batch", "peat batch", "run batch", "script"],
    "info.news": ["news", "newz", "nuze", "nuz", "latest news", "headlines"],
    "info.weather": ["weather", "whether", "wether", "forecast", "climate"],
    "time": ["time", "thyme", "tim", "what time", "current time"],
    "date": ["date", "dat", "day", "what date", "current date"],
    "now": ["now", "present", "current", "right now"],
    "delay": ["delay", "wait", "pause", "hold on"],
    "link": ["link", "open link", "go to", "navigate to", "open url"],
    "print": ["print", "say", "show", "display"],
    "quote": ["quote", "say quote", "random quote", "inspire me"],
    "debug": ["debug", "developer mode", "dev mode", "toggle debug"],
    "log": ["log", "logger", "write log"],
    "config.get": ["config.get", "get config", "config get", "read config"],
    "config.set": ["config.set", "set config", "save config", "config save"],
    "quit": ["quit", "exit", "bye", "goodbye", "close", "stop"],
    "cam": ["cam", "camera", "webcam", "open camera"],
}


# == Ensure Directories Exist ==
os.makedirs(camera_print_path, exist_ok=True)
os.makedirs(log_dump_folder, exist_ok=True)
os.makedirs(config_folder, exist_ok=True)
os.makedirs(EXT_UNACTIVE, exist_ok=True)
os.makedirs(EXT_ACTIVE, exist_ok=True)

class PEATAPI:
    def __init__(self, extension_id):
        self.router = router
        self.extension_id = extension_id
        self.ext = ExtMan()

    def register_command(self, name, handler):
        register_command(self.extension_id, name, handler)

    def register_help(self, help_dict):
        register_help(self.extension_id, help_dict)

    def get_commands(self):
        return list(router.keys())

    def print(self, text):
        print(text)

    def log(self, message, type="Info"):
        log(f"[EXT] [{self.extension_id}] {message}", "type")

    def set_camera_mode(self, m):
        set_camera_mode(m)

    def set_camera_index(self, i):
        set_camera_index(i)

    def version(self):
        return PEAT_VERSION

    def is_debug(self):
        return debug_mode

class ExtMan:
    def __init__(self):
        self.unactive = EXT_UNACTIVE
        self.active = EXT_ACTIVE

    def load(self, name):
        return ext_load(self, name)

    def unload(self, name):
        return ext_unload(self, name)

# ====== Definitions ======
# == System Functions ==
def init_log():
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("====== PEAT LOG ======\n")
        f.write(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("======================\n\n")

init_log()

def log(message, type="Info"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] [{type}] {message}\n"

    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line)

def log_dump():
    source = log_file

    # If no log exists, nothing to dump
    if not os.path.exists(source):
        print("No log file to dump.")
        return

    # Create timestamp
    timestamp = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")

    # Log the log was dumped (on the log)
    log(f"Log dumped", "Info")

    # Build new filename
    new_name = f"peat_log_{timestamp}.txt"

    # Copy file
    new_path = os.path.join(log_dump_folder, new_name)
    shutil.copy2(source, new_path)

    print(f"Log dumped to: {new_path}")

def crash_dump():
    log("Crash dump!!", "Info")
    log_dump()

def init_act_log():
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("====== PEAT ACTION LOG ======\n")
        f.write("What PEAT does behind the scenes...\n")
        f.write("(log for peat's automated actions)")
        f.write("=============================\n\n")

def shutdown(reason):
    print(f"PEAT encountered a critical error and is shutting down.")
    log(f"PEAT encountered a critical error and is shutting down.", "Warn")
    log(reason, "Error")
    crash_dump()
    sys.exit(1)

def quit(exit_code, who, reason):
    log(f"{who} flagged exit ({exit_code}) because: {reason}", "Info")
    print("Quitting PEAT...")
    if exit_code != 0:
        print("Error(s) were encountered. Check the log for more details.")
        if debug_mode:
            print("Debug mode is enabled. The log was automatically dumped for you.")
            log_dump()
    sys.exit(exit_code)

def log_clear():
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("")

def flag_error(who, message):
    log(f"Error flagged by {who} with message: '{message}'", "Error")

def validate_config_value(value_type, raw_value, default_value):

    if value_type == "str":
        if isinstance(raw_value, str) and raw_value.strip() != "":
            return raw_value
        return default_value

    if value_type == "int":
        if isinstance(raw_value, int):
            return raw_value

        if isinstance(raw_value, str):
            try:
                return int(raw_value)
            except ValueError:
                return default_value

        return default_value

    if value_type == "float":
        if isinstance(raw_value, (int, float)):
            return float(raw_value)

        if isinstance(raw_value, str):
            try:
                return float(raw_value)
            except ValueError:
                return default_value

        return default_value

    if value_type == "bool":

        # already a real bool
        if type(raw_value) is bool:
            return raw_value

        # convert strings
        if isinstance(raw_value, str):

            normalized = raw_value.strip().lower()

            if normalized in ("true", "1", "yes", "on", "enabled"):
                return True

            if normalized in ("false", "0", "no", "off", "disabled"):
                return False

        # convert integers
        if isinstance(raw_value, int):
            return bool(raw_value)

        return default_value

    if value_type == "list":
        return raw_value if isinstance(raw_value, list) else default_value

    if value_type == "dict":
        return raw_value if isinstance(raw_value, dict) else default_value

    return raw_value if isinstance(raw_value, type(default_value)) else default_value

def load_json_value(path, value_type, key_name, default_value):

    if not os.path.isfile(path):
        save_json_value(path, value_type, key_name, default_value)
        return default_value

    try:
        with open(path, "r", encoding="utf-8") as config_file:
            config_data = json.load(config_file)
            if debug_info:
                print(f"Loaded '{key_name}' with '{config_data}'")
    except (json.JSONDecodeError, OSError):
        save_json_value(path, value_type, key_name, default_value)
        return default_value

    if not isinstance(config_data, dict):
        save_json_value(path, value_type, key_name, default_value)
        return default_value

    raw_value = config_data.get(key_name, default_value)
    value = validate_config_value(value_type, raw_value, default_value)

    if key_name not in config_data or value != raw_value:
        config_data[key_name] = value
        try:
            with open(path, "w", encoding="utf-8") as config_file:
                json.dump(config_data, config_file, indent=2)
        except OSError:
            print("Warning: could not write config file.")

    return value

def save_json_value(path, value_type, key_name, save_value):

    config_data = {}
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as config_file:
                existing = json.load(config_file)
                if isinstance(existing, dict):
                    config_data = existing
        except (json.JSONDecodeError, OSError):
            config_data = {}

    config_data[key_name] = validate_config_value(value_type, save_value, save_value)

    try:
        with open(path, "w", encoding="utf-8") as config_file:
            json.dump(config_data, config_file, indent=2)
    except OSError:
        print("Warning: could not write config file.")

def make_dir(path, name, simple_path):
    new_folder = name

    if simple_path != (""):
        target_path = os.path.join(HOME, "Pictures", new_folder)
    else:
        target_path = os.path.join(path, new_folder)

    os.makedirs(target_path, exist_ok=True)
    print(f"Directory created at: {target_path}")
    log(f"Directory created at: {target_path}", "Info")

def load_json_data():

    log("Attempting to load JSON data...", "Info")

    try:
        global debug_info, name, dev_mode, debug_mode, tts_toggle, user_info_path
        global PEAT_VERSION

        # Load config data
        debug_info = load_json_value(config_path, "bool", "debugInfo", False)
        name = load_json_value(config_path, "str", "name", default_name)
        dev_mode = load_json_value(config_path, "bool", "dev", False)
        debug_mode = load_json_value(config_path, "bool", "debug", False)
        tts_toggle = load_json_value(config_path, "bool", "tts", False)

        # Load system data
        PEAT_VERSION = load_json_value(sys_info_path, "str", "sysVersion", "Error!!")

        
    except (Exception) as e:
        shutdown(f"Error while loading JSON data: {e}")

def register_command(extension, command_name, handler):
    global router

    if extension == "core":
        full_name = command_name
    else:
        full_name = f"{extension}.{command_name}"

    router[full_name] = handler

# == Extension Manifest System ==
ext_manifest_cache = {}

def get_ext_data(ext_id, item=None):
    """
    Get extension data from manifest.json.
    
    Args:
        ext_id: Extension folder name (e.g., 'cam', 'ai', 'info_fetch')
        item: Specific field to retrieve (e.g., 'namespace', 'name', 'version')
              If None, returns the entire manifest dict
    
    Returns:
        Requested item value, entire manifest dict, or None if not found
    """
    global ext_manifest_cache
    
    # Return from cache if available
    if ext_id in ext_manifest_cache:
        manifest = ext_manifest_cache[ext_id]
    else:
        # Load manifest from file
        manifest_path = os.path.join(EXT_ACTIVE, ext_id, "manifest.json")
        
        if not os.path.exists(manifest_path):
            return None
        
        try:
            with open(manifest_path, "r", encoding="utf-8-sig") as f:
                manifest = json.load(f)
                ext_manifest_cache[ext_id] = manifest
        except (json.JSONDecodeError, OSError) as e:
            log(f"[EXT] Error loading manifest for {ext_id}: {e}", "Error")
            return None
    
    # Return specific item or entire manifest
    if item is None:
        return manifest
    else:
        return manifest.get(item)

def clear_ext_manifest_cache():
    """Clear the extension manifest cache (call when reloading extensions)."""
    global ext_manifest_cache
    ext_manifest_cache.clear()

# == Extension Functions ==
def ext_load(self, name):
    name = name.replace(".py", "").strip()

    src_dir = os.path.join(self.unactive, name)
    dst_dir = os.path.join(self.active, name)

    if not os.path.isdir(src_dir):
        print(f"Extension '{name}' not found in Unactive.")
        return

    if os.path.isdir(dst_dir):
        print(f"Extension '{name}' is already active.")
        return

    shutil.move(src_dir, dst_dir)

    print(f"Extension '{name}' loaded.")
    log(f"[EXT] Loaded {name}", "Info")

    reload_extensions()

def ext_unload(self, name):
    name = name.replace(".py", "").strip()

    src_dir = os.path.join(self.active, name)
    dst_dir = os.path.join(self.unactive, name)

    if not os.path.isdir(src_dir):
        print(f"Extension '{name}' is not active.")
        return

    if os.path.isdir(dst_dir):
        print(f"Extension '{name}' is already unloaded.")
        return

    shutil.move(src_dir, dst_dir)

    print(f"Extension '{name}' unloaded.")
    log(f"[EXT] Unloaded {name}", "Info")

    reload_extensions()

def reload_extensions():
    global extensions, loaded_extensions

    log("Reloading extensions!", "Info")
    clear_ext_manifest_cache()
    loaded_extensions.clear()

    ext_folder = EXT_ACTIVE

    if not os.path.exists(ext_folder):
        os.makedirs(ext_folder, exist_ok=True)
        return

    # =========================
    # CLEAR OLD MODULES
    # =========================
    for name in list(extensions.keys()):
        if name in sys.modules:
            del sys.modules[name]
    extensions.clear()

    # =========================
    # LOAD EXTENSIONS
    # =========================
    for entry in os.listdir(ext_folder):
        ext_dir = os.path.join(ext_folder, entry)
        if not os.path.isdir(ext_dir):
            continue

        # Load manifest data
        manifest_path = os.path.join(ext_dir, "manifest.json")
        if not os.path.exists(manifest_path):
            log(f"[EXT] Skipping {entry}: manifest.json not found", "Warn")
            continue

        try:
            with open(manifest_path, "r", encoding="utf-8-sig") as handle:
                manifest = json.load(handle)
                ext_manifest_cache[entry] = manifest
        except (json.JSONDecodeError, OSError) as e:
            log(f"[EXT] Skipping {entry}: invalid manifest.json - {e}", "Error")
            continue

        # Get values from manifest
        module_file = manifest.get("module", f"{entry}.py")
        module_name = manifest.get("name", entry)
        extension_id = manifest.get("ext_id", entry)
        namespace = manifest.get("namespace", extension_id)

        module_path = os.path.join(ext_dir, module_file)
        if not os.path.exists(module_path):
            log(f"[EXT] Skipping {entry}: module not found at {module_path}", "Warn")
            continue

        spec = importlib.util.spec_from_file_location(module_name, module_path)
        if spec is None or spec.loader is None:
            log(f"[EXT] Could not create import spec for {entry}", "Error")
            continue

        module = importlib.util.module_from_spec(spec)

        # Inject API with namespace from manifest
        module.peat = PEATAPI(namespace)

        try:
            spec.loader.exec_module(module)

            if hasattr(module, "load_extension"):
                module.load_extension()

            extensions[module_name] = module
            loaded_extensions.append({
                "name": module_name,
                "namespace": namespace,
                "ext_id": extension_id,
                "version": manifest.get("version", "unknown")
            })

            log(f"[EXT] Loaded extension: {module_name} (namespace: {namespace})", "OK")

        except Exception as e:
            log(f"[EXT] {entry}: {e}", "Error")

def register_help(extension_name, help_dict):
    global extension_help
    extension_help[extension_name] = help_dict

# Extension API Functions
def test_extension_requirements(ext_id):
    requirements = get_ext_data(ext_id, "requirements")
    if requirements is None:
        return []
    
    if requirements <= loaded_extensions:
        return True
    else:
        pass


# == Extension Globals == 
def set_camera_mode(value):
    global camera_mode
    camera_mode = value

def set_camera_index(i):
    global camera_index
    camera_index = i

# == Command Handler Functions ==
def callback(indata, frames, time, status):
    if status:
        print(status)
    q.put(bytes(indata))

# == While True loop Functions ==
def camera_loop():
    global camera_mode

    cap = cv2.VideoCapture(camera_index)

    if not cap.isOpened():
        print(f"Failed to open camera {camera_index}.")
        camera_mode = False
        return

    previous_frame = None

    print("Q = Quit")
    print("P = Capture image")

    while camera_mode:

        ret, frame = cap.read()

        if not ret:
            print("Camera read failed.")
            break

        frame = cv2.resize(frame, (640, 360))

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (21, 21), 0)

        if previous_frame is None:
            previous_frame = gray
            continue

        delta = cv2.absdiff(previous_frame, gray)

        thresh = cv2.threshold(delta, 25, 255, cv2.THRESH_BINARY)[1]
        thresh = cv2.dilate(thresh, None, iterations=2)

        contours, _ = cv2.findContours(
            thresh,
            cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE
        )

        for contour in contours:
            if cv2.contourArea(contour) < 1200:
                continue

            x, y, w, h = cv2.boundingRect(contour)

            cv2.rectangle(
                frame,
                (x, y),
                (x + w, y + h),
                (0, 255, 0),
                2
            )

        previous_frame = gray

        cv2.imshow(
            f"PEAT Camera {camera_index}",
            frame
        )

        key = cv2.waitKey(1) & 0xFF

        handler = cam_controls.get(key)

        if handler:
            result = handler({"frame": frame})

            if result == "quit":
                break

    cap.release()
    cv2.destroyAllWindows()

# Other Functions
def get_time_period():
    # This was made by ai, but edited by me
    time_period = ""

    # 1. Get current time in Central Time (US/Central)
    now_ct = datetime.now()
    hour = now_ct.hour

    # 2. Categorize based on standard ranges
    if 5 <= hour < 12:
        time_period =  "morning"
    elif 12 <= hour < 17:
        time_period = "afternoon"
    elif 17 <= hour < 21:
        time_period ="evening"
    else:
        time_period = "night"

    return time_period

# Mem Functions
def load_user_info():
    global user_info

    if not os.path.isfile(user_info_path):
        save_user_info()
        return

    try:
        with open(user_info_path, "r", encoding="utf-8") as f:
            data = json.load(f)

            if isinstance(data, dict):
                user_info = data
            else:
                user_info = {}

    except Exception as e:
        log(f"Failed to load user info: {e}", "Error")
        user_info = {}

def save_user_info():
    try:
        with open(user_info_path, "w", encoding="utf-8") as f:
            json.dump(user_info, f, indent=2)

    except Exception as e:
        log(f"Failed to save user info: {e}", "Error")

# Do Functions
def get_tts_voices():
    voices = speaker.GetVoices()

    for i in range(voices.Count):
        print(voices.Item(i).GetDescription())

    speaker.Voice = voices.Item(0)

def suggest_command(cmd_act):
    if not router:
        return None

    keys = list(router.keys())

    # 1. exact prefix match (VERY important for dot commands)
    prefix_matches = [k for k in keys if k.startswith(cmd_act)]
    if prefix_matches:
        return prefix_matches[0]

    # 2. substring match (helps "weather" -> "info.weather")
    substring_matches = [k for k in keys if cmd_act in k]
    if substring_matches:
        return substring_matches[0]

    # 3. fuzzy match fallback
    matches = difflib.get_close_matches(cmd_act, keys, n=1, cutoff=0.6)
    if matches:
        return matches[0]

    return None

def parse_input(user_input):

    try:
        parts = shlex.split(user_input)

    except ValueError as e:
        print(f"Parse error: {e}")
        return "", "", ""

    cmd_act = parts[0] if len(parts) > 0 else ""
    cmd_args1 = parts[1] if len(parts) > 1 else ""
    cmd_args2 = " ".join(parts[2:]) if len(parts) > 2 else ""

    return cmd_act, cmd_args1, cmd_args2

def resolve_voice_alias(command_text):

    return command_text

    normalized = command_text.strip().lower()
    if not normalized:
        return command_text

    for canonical, aliases in voice_aliases.items():
        sorted_aliases = sorted(aliases, key=lambda x: len(x), reverse=True)
        for alias in sorted_aliases:
            alias_text = alias.lower().strip()
            if normalized == alias_text:
                return canonical
            if normalized.startswith(alias_text + " "):
                rest = command_text.strip()[len(alias_text):].strip()
                return f"{canonical} {rest}"

    return command_text

def clean_args(args, quote):
    args = args.strip()

    if len(args) >= 2 and args[0] == quote and args[-1] == quote:
        return args[1:-1]  # remove first and last quote

    return args

def speak(text):
    global speaker
    
    try:
        speaker.Speak("", 3)
        speaker.Speak(text)

    except Exception as e:
        log(f"[TTS ERROR] Failed to speak: {e}", "Warn")


    print(text)

    if tts_toggle:
        speak(text)

# Command Functions old
def launch_app(app_name):
    try:
        if sys.platform.startswith('win'):
            os.startfile(app_name)
        elif sys.platform.startswith('darwin'):
            subprocess.Popen(['open', '-a', app_name])
        else:
            subprocess.Popen([app_name])
        print(f"Launching '{app_name}'...")
    except Exception as e:
        print(f"Failed to launch '{app_name}': {e}")
        log(f"Failed to launch '{app_name}': {e}", "Warn")

def execute(exec_name):
    if exec_name in executables:
        print(f"Executing '{exec_name}'...")
    else:
        print(f"Attempted to execute '{exec_name}', but nothing was found with that name.")

# Camera Controls
def cam_quit(ctx):
    global camera_mode
    camera_mode = False
    return "quit"

def cam_capture(ctx):
    frame = ctx["frame"]

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    filename = os.path.join(
        camera_print_path,
        f"capture_{timestamp}.png"
    )

    cv2.imwrite(filename, frame)
    log(f"[CAM C] Saved capture: {filename}", "Info")
    print(f"Saved capture: {filename}")

    return "capture"

# PBAT Functions
def repeat_overflow(value):
    print(f"Overflow: loop repeats {value} times which exceeds max of {MAX_REPEAT} times")
    flag_error("PBAT", "repeat overflow")

def label_overflow(value):
    print(f"Overflow: label length {value} exceeds max of {MAX_LABEL_LENGTH} characters")
    flag_error("PBAT", "label overflow")

def resolve_vars(text, vars):
    out = ""
    i = 0

    while i < len(text):
        ch = text[i]

        # ESCAPE SEQUENCE
        if ch == "\\":
            i += 1
            if i < len(text):
                out += text[i]
            i += 1
            continue

        # VARIABLE
        if ch == "$":
            i += 1
            var_name = ""

            while i < len(text) and (text[i].isalnum() or text[i] == "_"):
                var_name += text[i]
                i += 1

            out += str(vars.get(var_name, f"<undefined:{var_name}>"))
            continue

        out += ch
        i += 1

    return out

def eval_condition(a, op, b, vars_dict):
    # resolve variables if needed
    a = vars_dict.get(a, a)
    b = vars_dict.get(b, b)

    try:
        a = int(a)
        b = int(b)
    except:
        pass  # keep as string if not numeric

    if op == "==":
        return a == b
    elif op == "!=":
        return a != b
    elif op == ">":
        return a > b
    elif op == "<":
        return a < b

    return False

def preprocess_pbat(lines):
    labels = {}
    cleaned = []

    for i, line in enumerate(lines):
        line = line.strip()

        if not line or line.startswith("*"):
            continue

        # LABEL
        if line.startswith("label "):
            name = line[6:].replace(":", "").strip()

            if len(name) > MAX_LABEL_LENGTH:
                label_overflow(len(name))
                continue

            labels[name] = len(cleaned)
            continue

        cleaned.append(line)

    return cleaned, labels

def run_pbat(script_name):

    if not script_name.endswith(".pbat"):
        script_name += ".pbat"

    script_path = os.path.join(pbat_trust_folder, script_name)

    if not os.path.exists(script_path):
        print(f"PBAT script not found: {script_name}")
        return

    print(f"Running PBAT: {script_name}")
    log(f"Running PBAT script: {script_name}", "Info")

    with open(script_path, "r", encoding="utf-8") as f:
        lines = [l.strip() for l in f.readlines()]

    vars_dict = {}
    labels = {}

    # =========================
    # LABEL PASS
    # =========================
    for i, line in enumerate(lines):
        if line.endswith(":") and not line.startswith("repeat") and not line.startswith("if"):
            label = line[:-1].strip()

            if len(label) > PBAT_MAX_LABEL_LEN:
                flag_error("PBAT", f"Overflow: label length {len(label)} exceeds max of {PBAT_MAX_LABEL_LEN}")
                return

            labels[label] = i

    # =========================
    # UTIL
    # =========================
    def resolve(text):
        text = text.replace("\\$", "__DOLLAR__")
        for k, v in vars_dict.items():
            text = text.replace(f"${k}", str(v))
        return text.replace("__DOLLAR__", "$")

    def find_block(start_idx):
        depth = 1
        i = start_idx
        while i < len(lines):
            if lines[i].startswith(("if", "repeat")):
                depth += 1
            elif lines[i] == "end":
                depth -= 1
                if depth == 0:
                    return i
            i += 1
        return -1

    # =========================
    # EXECUTION
    # =========================
    i = 0
    while i < len(lines):
        line = lines[i]

        if not line or line.startswith("*"):
            i += 1
            continue

        # skip labels
        if line.endswith(":"):
            i += 1
            continue

        parts = line.split(" ", 1)
        cmd = parts[0]
        args = parts[1] if len(parts) > 1 else ""

        # -------------------------
        # SET
        # -------------------------
        if cmd == "set":
            try:
                k, v = args.split(":", 1)
                vars_dict[k.strip()] = v.strip()
            except:
                flag_error("PBAT", "Syntax error in set" )
            i += 1
            continue

        # -------------------------
        # PRINT
        # -------------------------
        if cmd == "print":
            if args.startswith('"') and args.endswith('"'):
                raw = args[1:-1]
                print(resolve(raw))
            else:
                flag_error("PBAT", "Syntax error in print")
            i += 1
            continue

        # -------------------------
        # GOTO
        # -------------------------
        if cmd == "goto":
            target = args.strip()
            if target not in labels:
                flag_error("PBAT", f"Unknown label {target}")
                return
            i = labels[target]
            continue

        # -------------------------
        # REPEAT
        # -------------------------
        if cmd == "repeat":
            try:
                count = int(args.replace(":", "").strip())
            except:
                flag_error("PBAT", "Syntax error in repeat")
                i += 1
                continue

            if count > PBAT_MAX_REPEAT:
                flag_error("PBAT", f"Overflow: loop repeats {count} times which exceeds max of {PBAT_MAX_REPEAT} times")
                return

            start = i + 1
            end = find_block(start)

            if end == -1:
                flag_error("PBAT", "Missing end for repeat")
                return

            block = lines[start:end]

            for _ in range(count):
                j = 0
                while j < len(block):
                    sub = block[j]

                    if sub.startswith("set"):
                        try:
                            k, v = sub.split(" ", 1)[1].split(":", 1)
                            vars_dict[k.strip()] = v.strip()
                        except:
                            flag_error("PBAT", "Syntax error in repeat-set")

                    elif sub.startswith("print"):
                        try:
                            raw = sub.split(" ", 1)[1]
                            if raw.startswith('"') and raw.endswith('"'):
                                print(resolve(raw[1:-1]))
                        except:
                            pass

                    j += 1

            i = end + 1
            continue

        # -------------------------
        # IF
        # -------------------------
        if cmd == "if":
            try:
                cond = args.replace(":", "").split()
                if len(cond) != 3:
                    flag_error("PBAT", "Syntax error in if")
                    i += 1
                    continue

                a, op, b = cond

                a = vars_dict.get(a, a)
                b = vars_dict.get(b, b)

                try:
                    a = int(a)
                    b = int(b)
                except:
                    pass

                ok = (
                    a == b if op == "=="
                    else a != b if op == "!="
                    else a > b if op == ">"
                    else a < b if op == "<"
                    else False
                )

                end = find_block(i + 1)
                if end == -1:
                    flag_error("PBAT", "Missing end for if")
                    return

                if ok:
                    sub = lines[i + 1:end]
                    for s in sub:
                        do(s, "pbat")

                i = end + 1
                continue

            except:
                flag_error("PBAT", "Syntax error in if")
                i += 1
                continue

        # -------------------------
        # FALLBACK
        # -------------------------
        do(line, "pbat")
        i += 1

    log(f"Finished PBAT script: {script_name}", "OK")
    print(f"PBAT Script '{script_name}' finished.")

# Core commands
def cmd_help(a1, a2, who):
    if not a1:
        print("Commands:")

        print(f"\n[core]")
        for cmd in router:
            print(f"- {cmd}")

        for ext, cmds in extension_help.items():
            print(f"\n[{ext}]")
            for k, v in cmds.items():
                print(f"- {k}: {v}")

        return

def cmd_log(a1, a2, who):
    if who == "pbat":
        log(f"[PBAT] {a1}", "Devlog")
        return

    if not debug_mode:
        print("This feature is for developers only. Enable debug mode to use it.")
        return

    # log dump
    if a1 == "dump":
        log_dump()
        return

    # log clear
    if a1 == "clear":
        log_clear()
        print("Log cleared.")
        return

    # write custom log message
    if a1 and a1[0] in ('"', "'"):
        a1 = clean_args(a1, a1[0])
        log(a1, "Devlog")
        print("Logged message.")
        return

    print("Unknown log command.")

def cmd_launch(a1, a2, who):
    if not a1:
        print("Expected a string for the application name, but got nothing.")
        return

    if a1[0] in ('"', "'"):
        a1 = clean_args(a1, a1[0])

    launch_app(a1)

def cmd_exec(a1, a2, who):
    if who != "pbat":
        execute(a1)
    else:
        print("Cannot execute from within a PBAT script!")
        flag_error("EXEC", "Attempted to execute from within a PBAT script.")

def cmd_pbat(a1, a2, who):
    if who == "pbat":
        print("Cannot run a PBAT script from within another PBAT script!")
        flag_error("PBAT", "Attempted nested PBAT.")
        return

    if not a1:
        print("Expected a string for the script name, but got nothing.")
        return

    if a1[0] in ('"', "'"):
        a1 = clean_args(a1, a1[0])

    run_pbat(a1)

def cmd_print(a1, a2, who):
    if not a1:
        print("Expected a string to print, but got nothing.")
        return

    if a1[0] in ('"', "'"):
        a1 = clean_args(a1, a1[0])

    print(a1)

def cmd_quote(a1, a2, who):
    print(quotes[total_commands_used % len(quotes)])

def cmd_link(a1, a2, who):
    if not a1:
        print("Expected a string for the URL, but got nothing.")
        return

    if a1[0] in ('"', "'"):
        a1 = clean_args(a1, a1[0])

    print(f"Opening link in browser: {a1}")
    webbrowser.open(a1)

def cmd_delay(a1, a2, who):
    if not a1:
        print("Expected a number, but got nothing.")
        return

    try:
        delay_time = float(a1)
        time.sleep(delay_time)
        print(f"Delayed for {delay_time} seconds.")
    except ValueError:
        flag_error("delay", f"Expected a number for delay time, got: {a1}")

def cmd_time(a1, a2, who):
    now = datetime.now()
    print(f"Current time: {now.strftime('%H:%M:%S')}")

def cmd_date(a1, a2, who):
    now = datetime.now()
    print(f"Today is: {now.strftime('%Y-%m-%d')}")

def cmd_now(a1, a2, who):
    now = datetime.now()
    print(f"Date: {now.strftime('%Y-%m-%d')}")
    print(f"Time: {now.strftime('%H:%M:%S')}")

def cmd_debug(a1, a2, who):
    global debug_mode

    debug_mode = not debug_mode

    print(f"Debug mode is {'enabled' if debug_mode else 'disabled'}.")
    log(f"Debug mode {'enabled' if debug_mode else 'disabled'} by user.", "Info")

def cmd_config_get(a1, a2, who):
    if not a1:
        print("Expected config key.")
        return

    value = load_json_value(config_path, "str", a1, "undefined")
    print(f"{a1} = {value}")

def cmd_config_set(a1, a2, who):
    if not a1 or not a2:
        print("Usage: config.set <key> <value>")
        return

    # detect value type automatically
    value = a2
    lowered = value.lower()

    if lowered == "true":
        value = True
    elif lowered == "false":
        value = False
    else:
        try:
            value = int(value)
        except:
            try:
                value = float(value)
            except:
                pass

    save_json_value(config_path, type(value).__name__, a1, value)
    print(f"Saved: {a1} = {value}")

def cmd_version(a1, a2, who):
    print(f"You are using P.E.A.T. version {PEAT_VERSION}")

# Extman commands
def cmd_ext_load(a1, a2, who):
    api.ext.load(a1)

def cmd_ext_unload(a1, a2, who):
    api.ext.unload(a1)

def cmd_ext_reload(a1, a2, who):
    reload_extensions()

def cmd_ext_list(a1, a2, who):
    if not loaded_extensions:
        print("No extensions are currently loaded.")
        return
    
    print("Currently loaded extensions:")
    for i, ext in enumerate(loaded_extensions, 1):
        print(f"\n{i}. {ext['name']}")
        print(f"   Namespace: {ext['namespace']}")
        print(f"   ID: {ext['ext_id']}")
        print(f"   Version: {ext['version']}")

# Dev commands
def devcmd_test_dev(a1, a2, who):
    print("The dev test command works!")

def devcmd_self_check(a1, a2, who):
    pass

def devcmd_get_calue(a1, a2, who):
    pass


# == Populate Dictionaries == 
def populate_router():
    global router

    register_command("core", "help", cmd_help)
    register_command("core", "v", cmd_version)

    # ExtMan commands
    # register_command("core", "ext.get", cmd_ext_get)
    register_command("core", "ext.load", cmd_ext_load)
    register_command("core", "ext.unload", cmd_ext_unload)
    register_command("core", "ext.reload", cmd_ext_reload)
    register_command("core", "ext.list", cmd_ext_list)

    # Dev commands + debug commands
    register_command("core", "dev.test", devcmd_test_dev)
    register_command("core", "dev.forceshut", lambda a1, a2, who: shutdown("Forced shutdown by user: dev. Good job!"))
    register_command("core", "dev.check", devcmd_self_check)

    # Action commands
    register_command("act", "launch", cmd_launch)
    register_command("core", "exec", cmd_exec)
    register_command("core", "pbat", cmd_pbat)

    # Time and date commands
    register_command("core", "time", cmd_time)
    register_command("core", "date", cmd_date)
    register_command("core", "now", cmd_now)

    # Debug + developer commands
    register_command("core", "debug", cmd_debug)
    register_command("core", "log", cmd_log)

    # Database commands
    register_command("core", "getcfg", cmd_config_get)
    register_command("core", "setcfg", cmd_config_set)

    # PBAT commands (usually)
    register_command("core", "delay", cmd_delay)
    register_command("core", "link", cmd_link)
    register_command("core", "print", cmd_print)

    # Misc commands
    register_command("core", "quote", cmd_quote)

def populate_cam_controls():
    global cam_controls

    cam_controls[ord("q")] = cam_quit
    cam_controls[ord("p")] = cam_capture

api = PEATAPI("core")
populate_router()
reload_extensions()
populate_cam_controls()

# The main do function
def do(cmd, who):
    log(f"Attempting do: '{cmd}'", "Info")
    if who != "user":
        log(f"Command issued by: {who}, this will be logged in action log.", "Info")
    global total_commands_used
    total_commands_used += 1

    cmd_act, cmd_args1, cmd_args2 = parse_input(cmd)

    if False:
        original_cmd = cmd
        cmd = resolve_voice_alias(cmd)
        if debug_mode and original_cmd != cmd:
            print(f"[ALIAS] '{original_cmd}' -> '{cmd}'")

    cmd_act, cmd_args1, cmd_args2 = parse_input(cmd)

    if cmd_act == "end":
        if who == "pbat":
            return

    if cmd.lower() == fallback_input.lower():
        flag_error("DO", "Fallback input triggered")
        return
    
    # hardcoded protected commands
    if cmd_act in ("q", "quit", "exit", "bye"):

        if who == "user":
            quit(0, "User", "User requested exit.")

        elif who == "pbat":
            print("Cannot quit PEAT from within a PBAT script!")
            flag_error("PBAT", "Attempted quit inside PBAT.")

        elif who == "sys":
            quit(0, "System", "System requested exit.")

        return

    # router dispatch
    handler = router.get(cmd_act)

    if handler:
        try:
            handler(cmd_args1, cmd_args2, who)
        except Exception as e:
            print(f"Command error: {e}")
            log(f"Command error in '{cmd_act}': {e}", "Error")
    else:
        suggestion = suggest_command(cmd_act)

        print(f"Unknown command: '{cmd_act}'")
        log(f"Unknown command: '{cmd_act}'", "Warn")

        if suggestion:
            print(f"Did you mean: '{suggestion}'?")

    print()

# == NEW Load Config from JSON ==
load_json_data()
load_user_info()

if debug_mode:
    log("[INIT] Developer mode enabled via config.", "Info")
    print("Developer mode is enabled. Please actually do something!\n")


log(f"[INIT] Dictionary 'router' populated with {len(router)} commands.", "OK")
log(f"[INIT] Dictionary 'cam_controls' populated with {len(cam_controls)} commands.", "OK")
log("[INIT] Config from JSON fully loaded.", "OK")
log("[INIT] PEAT initialization complete.", "DONE")

if get_time_period() == "night":
    print(f"It's pretty late, {name}.")
else:
    print(f"Good {get_time_period()}, {name}.")

print(query_placeholder)

log("====== ENTERING MAIN LOOP ======")

while True:
    try:
        if camera_mode:
            camera_loop()
            continue

        # 2. INPUT SOURCE SWITCH
        if False:
            query = listen()
            if query and query.strip():
                if voice_mode:
                    print(f"$v {query}")
                    do(query, "user")
                continue

        else:
            query = input("$ ")

        # ignore empty input
        if not query or query.strip() == "":
            continue

        # print(f"$ {query}")

        # 3. Execute command
        do(query, "user")

        time.sleep(0.01)

    except KeyboardInterrupt:
            print("KeyboardInterrupt")
            log("KeyboardInterrupt: Exiting main loop.", "Info")
            quit(0, "User", "User KeyboardInterrupt.")

    except Exception as e:
        shutdown(f"Main loop: {e}")
