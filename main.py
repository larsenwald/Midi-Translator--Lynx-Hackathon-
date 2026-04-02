import os
import sys
import json
import threading
import time
import mido
import webview
import pystray
from PIL import Image, ImageDraw
import win32gui
import win32con

# ─── PATHS ────────────────────────────────────────────────────────────────────

APP_NAME         = "MidiWarp"
OUTPUT_PORT_NAME = "MidiWarp_OUT"  # Your loopMIDI port name — change if needed

DATA_DIR    = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
SCRIPTS_DIR = os.path.join(DATA_DIR, "scripts")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
os.makedirs(SCRIPTS_DIR, exist_ok=True)

try:
    from google import genai as google_genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False
    print("[AI] google-genai not installed. Run: pip install google-genai")

AI_SYSTEM_PROMPT = """\
You are an expert at writing MidiWarp scripts.
MidiWarp scripts are plain Python files that run on every incoming MIDI event.

AVAILABLE VARIABLES (read/write):
  event_type : str  — 'note_on', 'note_off', 'control_change', 'pitchwheel', 'program_change'
  channel    : int  — 1-16
  note       : int  — 0-127  (note_on / note_off only)
  velocity   : int  — 0-127  (note_on / note_off only)
  cc_num     : int  — 0-127  (control_change only)
  cc_value   : int  — 0-127  (control_change only)
  pitch      : int  — -8192 to +8191  (pitchwheel only)
  program    : int  — 0-127  (program_change only)

NOTE NAME CONSTANTS (all available directly, e.g. C4=60, Cs4=61, Db4=61, D4=62 ... up to G9=127):
  Natural notes: C0, D0, E0, F0, G0, A0, B0, C1 ... etc
  Sharps: Cs4 = C#4, Ds4 = D#4, Fs4 = F#4, Gs4 = G#4, As4 = A#4
  Flats:  Db4 = Db4, Eb4, Gb4, Ab4, Bb4  (enharmonic equivalents of sharps)

HELPER FUNCTIONS:
  send_note_on(note, velocity, channel=None)
  send_note_off(note, velocity=0, channel=None)
  send_cc(cc_num, value, channel=None)
  send_pitchbend(pitch, channel=None)   — pitch: -8192 to +8191
  send_program_change(program, channel=None)
  block()  — suppress the original event entirely

RULES:
- Mutating a variable (e.g. note = 74) rewrites the outgoing message.
- Calling send_*() emits an additional message alongside the original.
- Calling block() suppresses the original. Combine with send_*() to replace it.
- Only access variables that exist for the current event_type. Always guard with if event_type == '...' first.
- Scripts are stateless — no data persists between events.
- Do not import anything. Do not use output(). Do not use raw status bytes.
- Always guard by event_type before accessing event-specific variables.

OUTPUT FORMAT:
Return ONLY the raw Python script code, no markdown fences, no explanation, no comments unless they aid clarity.
"""


def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"[Config] Save error: {e}")


# ─── MIDI HANDLER ─────────────────────────────────────────────────────────────

def strip_port_number(name):
    """'Impact GX61 0' -> 'Impact GX61'"""
    parts = name.rsplit(' ', 1)
    if len(parts) == 2 and parts[1].isdigit():
        return parts[0]
    return name

def find_port_by_substring(names, substring):
    """Return the first port whose name contains substring, or None."""
    for n in names:
        if substring in n:
            return n
    return None


class MidiHandler:
    def __init__(self):
        self.input_port  = None
        self.output_port = None
        self._lock       = threading.Lock()

    def get_input_names(self):
        """All input ports except our own loopMIDI output.
        Returns list of {"full": "Impact GX61 0", "display": "Impact GX61"}.
        """
        try:
            ports = [n for n in mido.get_input_names() if OUTPUT_PORT_NAME not in n]
            return [{"full": n, "display": strip_port_number(n)} for n in ports]
        except Exception:
            return []

    def set_input(self, full_name):
        with self._lock:
            if self.input_port:
                try:
                    self.input_port.close()
                except Exception:
                    pass
                self.input_port = None
            try:
                self.input_port = mido.open_input(full_name)
                print(f"[MIDI] Input: {full_name}")
                return True
            except Exception as e:
                print(f"[MIDI] Input error: {e}")
                return False

    def open_output(self):
        with self._lock:
            if self.output_port:
                return True
            try:
                all_outputs = mido.get_output_names()
                full_name = find_port_by_substring(all_outputs, OUTPUT_PORT_NAME)
                if not full_name:
                    print(f"[MIDI] Output port containing '{OUTPUT_PORT_NAME}' not found")
                    return False
                self.output_port = mido.open_output(full_name)
                print(f"[MIDI] Output: {full_name}")
                return True
            except Exception as e:
                print(f"[MIDI] Output error: {e}")
                return False

    def send(self, msg):
        with self._lock:
            if self.output_port:
                try:
                    self.output_port.send(msg)
                except Exception as e:
                    print(f"[MIDI] Send error: {e}")

    def iter_pending(self):
        with self._lock:
            port = self.input_port
        if port:
            try:
                yield from port.iter_pending()
            except Exception as e:
                # Port died — clear it and signal disconnection to caller
                with self._lock:
                    self.input_port = None
                raise ConnectionError(f"Input port disconnected: {e}")

    def close(self):
        with self._lock:
            for port in (self.input_port, self.output_port):
                if port:
                    try:
                        port.close()
                    except Exception:
                        pass
            self.input_port  = None
            self.output_port = None


# ─── SCRIPT ENGINE ────────────────────────────────────────────────────────────

NEW_SCRIPT_TEMPLATE = """\
# event_type : 'note_on', 'note_off', 'control_change', 'pitchwheel', 'program_change'
# channel    : 1-16  (read/write)
#
# note_on / note_off:
#   note      : 0-127  (read/write)   e.g. 60 = C4, 72 = C5
#   velocity  : 0-127  (read/write)
#
# control_change:
#   cc_num    : 0-127  (read/write)
#   cc_value  : 0-127  (read/write)
#
# pitchwheel:
#   pitch     : -8192 to +8191  (read/write)
#
# program_change:
#   program   : 0-127  (read/write)
#
# Helpers:
#   send_note_on(note, velocity, channel=None)
#   send_note_off(note, velocity=0, channel=None)
#   send_cc(cc_num, value, channel=None)
#   send_pitchbend(pitch, channel=None)
#   send_program_change(program, channel=None)
#   block()  -- suppress this event, don't forward it
#
# Mutating a variable rewrites the outgoing message.
# send_* emits an additional message alongside the original.
# block() suppresses the original; pair with send_* to replace it entirely.
# If you do nothing, the message passes through unchanged.

"""


class ScriptEngine:
    def __init__(self):
        self.scripts = []   # [{"name": str, "code": str, "enabled": bool}]

    # ── disk ops ──

    def load_scripts(self):
        self.scripts = []
        try:
            files = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith(".py")]
        except Exception:
            files = []
        cfg = load_config()
        enabled_states = cfg.get("script_enabled", {})
        order = cfg.get("script_order", [])

        # Sort files by saved order; unknowns appended alphabetically at the end
        ordered = [n + ".py" for n in order if (n + ".py") in files]
        remaining = sorted(f for f in files if f not in ordered)
        files = ordered + remaining

        for fname in files:
            path = os.path.join(SCRIPTS_DIR, fname)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    code = f.read()
                name = fname[:-3]
                self.scripts.append({
                    "name":       name,
                    "code":       code,
                    "enabled":    enabled_states.get(name, True),
                    "last_error": None,
                })
                print(f"[Scripts] Loaded: {fname}")
            except Exception as e:
                print(f"[Scripts] Error loading {fname}: {e}")

    def save_script(self, name, code):
        path = os.path.join(SCRIPTS_DIR, name + ".py")
        with open(path, "w", encoding="utf-8") as f:
            f.write(code)
        self.load_scripts()

    def delete_script(self, name):
        path = os.path.join(SCRIPTS_DIR, name + ".py")
        if os.path.exists(path):
            os.remove(path)
        self.load_scripts()

    def rename_script(self, old_name, new_name):
        old_path = os.path.join(SCRIPTS_DIR, old_name + ".py")
        new_path = os.path.join(SCRIPTS_DIR, new_name + ".py")
        if os.path.exists(old_path):
            os.rename(old_path, new_path)
        self.load_scripts()

    def set_enabled(self, name, enabled):
        for s in self.scripts:
            if s["name"] == name:
                s["enabled"] = enabled
                cfg = load_config()
                states = cfg.get("script_enabled", {})
                states[name] = enabled
                cfg["script_enabled"] = states
                save_config(cfg)
                return

    # ── processing ──

    def process_message(self, msg):
        """
        Run msg through all enabled scripts.
        Returns (output_msgs, applied_script_name_or_None).
        """
        enabled = [s for s in self.scripts if s["enabled"]]
        if not enabled:
            return [msg], None

        # Mutable context — scripts read and write these directly
        ctx = {
            "event_type": msg.type,
            "channel":    msg.channel + 1,  # 1-based for humans
        }
        if msg.type in ("note_on", "note_off"):
            ctx["note"]     = msg.note
            ctx["velocity"] = msg.velocity
        elif msg.type == "control_change":
            ctx["cc_num"]   = msg.control
            ctx["cc_value"] = msg.value
        elif msg.type == "pitchwheel":
            ctx["pitch"]    = msg.pitch
        elif msg.type == "program_change":
            ctx["program"]  = msg.program

        # Note name constants — C4=60, Cs4=61, Db4=61, etc.
        NOTE_NAMES_SHARP = ['C','Cs','D','Ds','E','F','Fs','G','Gs','A','As','B']
        NOTE_NAMES_FLAT  = ['C','Db','D','Eb','E','F','Gb','G','Ab','A','Bb','B']
        for midi_num in range(128):
            octave = (midi_num // 12) - 1
            ctx[f'{NOTE_NAMES_SHARP[midi_num % 12]}{octave}'] = midi_num
            ctx[f'{NOTE_NAMES_FLAT[midi_num % 12]}{octave}']  = midi_num

        extra_msgs = []
        blocked    = [False]
        applied    = [None]

        def _ch(channel):
            return max(0, min(15, int(channel) - 1))

        def send_note_on(note, velocity, channel=None):
            ch = _ch(channel if channel is not None else ctx["channel"])
            extra_msgs.append(mido.Message("note_on", channel=ch,
                                           note=int(note), velocity=int(velocity)))

        def send_note_off(note, velocity=0, channel=None):
            ch = _ch(channel if channel is not None else ctx["channel"])
            extra_msgs.append(mido.Message("note_off", channel=ch,
                                           note=int(note), velocity=int(velocity)))

        def send_cc(cc_num, value, channel=None):
            ch = _ch(channel if channel is not None else ctx["channel"])
            extra_msgs.append(mido.Message("control_change", channel=ch,
                                           control=int(cc_num),
                                           value=max(0, min(127, int(value)))))

        def send_pitchbend(pitch, channel=None):
            ch = _ch(channel if channel is not None else ctx["channel"])
            extra_msgs.append(mido.Message("pitchwheel", channel=ch,
                                           pitch=int(max(-8192, min(8191, pitch)))))

        def send_program_change(program, channel=None):
            ch = _ch(channel if channel is not None else ctx["channel"])
            extra_msgs.append(mido.Message("program_change", channel=ch,
                                           program=max(0, min(127, int(program)))))

        def block():
            blocked[0] = True

        ctx.update({
            "send_note_on":        send_note_on,
            "send_note_off":       send_note_off,
            "send_cc":             send_cc,
            "send_pitchbend":      send_pitchbend,
            "send_program_change": send_program_change,
            "block":               block,
        })

        for script in enabled:
            n_extra_before = len(extra_msgs)
            blocked_before = blocked[0]
            try:
                exec(script["code"], ctx)
            except Exception as e:
                import traceback
                err_str = traceback.format_exc().strip()
                script["last_error"] = err_str
                print(f"[Scripts] Error in '{script['name']}': {e}")
            else:
                script["last_error"] = None
            if (len(extra_msgs) > n_extra_before
                    or blocked[0] != blocked_before
                    or self._mutated(ctx, msg)):
                applied[0] = script["name"]

        output_msgs = []

        if not blocked[0]:
            try:
                ch = _ch(ctx["channel"])
                if msg.type in ("note_on", "note_off"):
                    rebuilt = mido.Message(msg.type, channel=ch,
                                           note=max(0, min(127, int(ctx["note"]))),
                                           velocity=max(0, min(127, int(ctx["velocity"]))))
                elif msg.type == "control_change":
                    rebuilt = mido.Message("control_change", channel=ch,
                                           control=max(0, min(127, int(ctx["cc_num"]))),
                                           value=max(0, min(127, int(ctx["cc_value"]))))
                elif msg.type == "pitchwheel":
                    rebuilt = mido.Message("pitchwheel", channel=ch,
                                           pitch=int(max(-8192, min(8191, ctx["pitch"]))))
                elif msg.type == "program_change":
                    rebuilt = mido.Message("program_change", channel=ch,
                                           program=max(0, min(127, int(ctx["program"]))))
                else:
                    rebuilt = msg
            except Exception as e:
                print(f"[Scripts] Rebuild error: {e}")
                rebuilt = msg
            output_msgs.append(rebuilt)

        output_msgs.extend(extra_msgs)

        # blocked with no send_* calls means nothing gets forwarded
        return output_msgs, applied[0]

    def _mutated(self, ctx, msg):
        """True if any context var differs from the original message."""
        if ctx.get("channel") != (msg.channel + 1):
            return True
        if msg.type in ("note_on", "note_off"):
            return ctx.get("note") != msg.note or ctx.get("velocity") != msg.velocity
        if msg.type == "control_change":
            return ctx.get("cc_num") != msg.control or ctx.get("cc_value") != msg.value
        if msg.type == "pitchwheel":
            return ctx.get("pitch") != msg.pitch
        if msg.type == "program_change":
            return ctx.get("program") != msg.program
        return False


# ─── WIN32 HELPERS ──────────────────────────────────────────────────────────────

def get_hwnd(window):
    """Get the native Win32 window handle from a pywebview window."""
    try:
        return window.gui.window.wid   # pywebview 4.x EdgeChromium
    except AttributeError:
        pass
    try:
        return win32gui.FindWindow(None, window.title)
    except Exception:
        return None

def hide_to_tray(hwnd):
    """Remove from taskbar by stripping WS_EX_APPWINDOW and adding WS_EX_TOOLWINDOW."""
    if not hwnd:
        return
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    ex_style &= ~win32con.WS_EX_APPWINDOW
    ex_style |=  win32con.WS_EX_TOOLWINDOW
    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)

def show_from_tray(hwnd):
    """Restore taskbar presence and bring window back."""
    if not hwnd:
        return
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    ex_style |=  win32con.WS_EX_APPWINDOW
    ex_style &= ~win32con.WS_EX_TOOLWINDOW
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

# ─── PYWEBVIEW API ────────────────────────────────────────────────────────────

class API:
    def __init__(self, midi: MidiHandler, engine: ScriptEngine):
        self.midi    = midi
        self.engine  = engine
        self._window = None

    def set_window(self, window):
        self._window = window

    def get_input_ports(self):
        return self.midi.get_input_names()

    def set_input(self, full_name):
        return self.midi.set_input(full_name)

    def get_output_port_name(self):
        return OUTPUT_PORT_NAME

    def get_scripts(self):
        return [{"name": s["name"], "code": s["code"], "enabled": s["enabled"],
                 "last_error": s.get("last_error")}
                for s in self.engine.scripts]

    def save_script(self, name, code):
        self.engine.save_script(name, code)
        return True

    def delete_script(self, name):
        self.engine.delete_script(name)
        return True

    def rename_script(self, old_name, new_name):
        self.engine.rename_script(old_name, new_name)
        return True

    def set_script_enabled(self, name, enabled):
        self.engine.set_enabled(name, enabled)
        return True

    def set_script_order(self, names):
        """Persist the script pipeline order and reorder in memory."""
        cfg = load_config()
        cfg["script_order"] = names
        save_config(cfg)
        # Reorder in-memory scripts to match
        order_map = {n: i for i, n in enumerate(names)}
        self.engine.scripts.sort(key=lambda s: order_map.get(s["name"], 9999))
        return True

    def get_api_key_set(self):
        cfg = load_config()
        return bool(cfg.get("gemini_api_key", ""))

    def set_api_key(self, key):
        cfg = load_config()
        cfg["gemini_api_key"] = key.strip()
        save_config(cfg)
        return True

    def generate_script(self, prompt):
        if not GENAI_AVAILABLE:
            return {"error": "google-genai not installed. Run: pip install google-genai"}
        cfg = load_config()
        api_key = cfg.get("gemini_api_key", "").strip()
        if not api_key:
            return {"error": "No API key set. Enter your Gemini API key first."}
        try:
            client = google_genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"{AI_SYSTEM_PROMPT}\n\nUser request: {prompt}",
            )
            code = response.text.strip()
            # Strip markdown fences if the model included them anyway
            if code.startswith("```"):
                lines = code.splitlines()
                code = "\n".join(
                    l for l in lines if not l.strip().startswith("```")
                ).strip()
            return {"code": code}
        except Exception as e:
            return {"error": str(e)}

    def move_window(self, x, y):
        if self._window:
            self._window.move(int(x), int(y))
            self._save_bounds(int(x), int(y), None, None)

    def resize_window(self, w, h, x, y):
        if self._window:
            self._window.resize(int(w), int(h))
            self._window.move(int(x), int(y))
            self._save_bounds(int(x), int(y), int(w), int(h))

    def _save_bounds(self, x, y, w, h):
        cfg    = load_config()
        bounds = cfg.get("window_bounds", {})
        if x is not None: bounds["x"] = x
        if y is not None: bounds["y"] = y
        if w is not None: bounds["w"] = w
        if h is not None: bounds["h"] = h
        cfg["window_bounds"] = bounds
        save_config(cfg)

    def minimize(self):
        if self._window: self._window.minimize()

    def hide_to_tray(self):
        """X button — hide window and remove from taskbar."""
        hwnd = get_hwnd(self._window)
        hide_to_tray(hwnd)

    def maximize(self):
        if self._window: self._window.maximize()

    def restore(self):
        if self._window: self._window.restore()


# ─── MIDI MONITOR THREAD ──────────────────────────────────────────────────────

NOTE_NAMES = ["C","C#","D","D#","E","F","F#","G","G#","A","A#","B"]

def note_name(n):
    return NOTE_NAMES[n % 12] + str((n // 12) - 1)

def msg_to_event(msg, applied):
    ev = {"type": msg.type, "applied": applied}
    if msg.type in ("note_on", "note_off"):
        ev.update({"note": msg.note, "noteName": note_name(msg.note),
                   "vel": msg.velocity, "ch": msg.channel + 1})
    elif msg.type == "control_change":
        ev.update({"cc": msg.control, "val": msg.value, "ch": msg.channel + 1})
    elif msg.type == "pitchwheel":
        ev.update({"pitch": msg.pitch, "ch": msg.channel + 1})
    elif msg.type == "program_change":
        ev.update({"program": msg.program, "ch": msg.channel + 1})
    else:
        ev.update({"raw": " ".join(f"{b:02X}" for b in msg.bytes())})
    return ev

def monitor_loop(midi: MidiHandler, engine: ScriptEngine, api: API):
    while True:
        try:
            for msg in midi.iter_pending():
                out_msgs, applied = engine.process_message(msg)

                for out_msg in out_msgs:
                    midi.send(out_msg)

                raw_ev = msg_to_event(msg, None)
                tx_ev  = msg_to_event(out_msgs[0], applied) if out_msgs else raw_ev

                script_errors = {
                    s["name"]: s["last_error"]
                    for s in engine.scripts
                    if s.get("last_error")
                }

                payload = json.dumps({
                    "raw": raw_ev,
                    "tx":  tx_ev,
                    "script_errors": script_errors,
                })

                if api._window:
                    try:
                        api._window.evaluate_js(f"window.onMidiEvent({payload})")
                    except Exception:
                        pass

        except ConnectionError as e:
            print(f"[MIDI] {e}")
            if api._window:
                try:
                    api._window.evaluate_js("window.onMidiDisconnected()")
                except Exception:
                    pass

        time.sleep(0.001)


def port_watchdog(midi: MidiHandler, api: API):
    """Poll the OS port list every second. If the connected port vanishes, signal JS."""
    while True:
        time.sleep(1)
        with midi._lock:
            port = midi.input_port
            connected_name = getattr(port, 'name', None) if port else None
        if connected_name:
            try:
                available = mido.get_input_names()
            except Exception:
                continue
            still_there = any(
                connected_name in n or n in connected_name
                for n in available
            )
            if not still_there:
                with midi._lock:
                    midi.input_port = None
                print(f"[MIDI] Port vanished: {connected_name}")
                if api._window:
                    try:
                        api._window.evaluate_js("window.onMidiDisconnected()")
                    except Exception:
                        pass




def build_tray_image(base_dir):
    """Load the tray icon from disk, falling back to a generated image."""
    icon_path = os.path.join(base_dir, "MidiWarp_tray.png")
    try:
        return Image.open(icon_path).convert("RGBA")
    except Exception:
        # Fallback: draw the old M shape
        img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
        dc  = ImageDraw.Draw(img)
        dc.ellipse([2, 2, 62, 62], fill=(30, 27, 22, 255))
        pts = [(14,46),(14,18),(32,36),(50,18),(50,46)]
        dc.line(pts, fill=(201, 169, 110, 255), width=5)
        return img


def setup_tray(window, midi, base_dir):
    """Create and run the system tray icon in a background thread."""

    def show(icon, item):
        hwnd = get_hwnd(window)
        show_from_tray(hwnd)

    def quit_app(icon, item):
        icon.stop()
        midi.close()
        window.destroy()

    icon = pystray.Icon(
        "MidiWarp",
        build_tray_image(base_dir),
        "MidiWarp",
        menu=pystray.Menu(
            pystray.MenuItem("Show", show, default=True),
            pystray.MenuItem("Quit", quit_app),
        )
    )

    t = threading.Thread(target=icon.run, daemon=True)
    t.start()
    return icon


# ─── STARTUP CHECKS ───────────────────────────────────────────────────────────

import subprocess
import winreg
import psutil

LOOPMIDI_PATHS = [
    r"C:\Program Files (x86)\Tobias Erichsen\loopMIDI\loopMIDI.exe",
    r"C:\Program Files\Tobias Erichsen\loopMIDI\loopMIDI.exe",
]

def _push(window, key, label, status, message=None):
    if not window:
        return
    msg_js = f', {json.dumps(message)}' if message else ', null'
    try:
        window.evaluate_js(
            f'window.onStartupStep({json.dumps(key)}, {json.dumps(label)}, {json.dumps(status)}{msg_js})'
        )
    except Exception:
        pass

def _find_loopmidi_exe():
    for path in LOOPMIDI_PATHS:
        if os.path.exists(path):
            return path
    return None

def run_startup_checks(window, base_dir):
    """Run pre-flight checks, pushing status to the splash screen as we go."""

    # ── Check 1: loopMIDI installed ──
    _push(window, 'loopmidi', 'Checking loopMIDI', 'active')
    exe_path = _find_loopmidi_exe()

    if exe_path:
        _push(window, 'loopmidi', 'loopMIDI found', 'done')
    else:
        # Attempt silent install from bundled setup
        installer = os.path.join(base_dir, 'loopMIDISetup.exe')
        if os.path.exists(installer):
            _push(window, 'loopmidi', 'Installing loopMIDI…', 'active')
            try:
                subprocess.call(
                    [installer, '/silent'],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                time.sleep(3)
                exe_path = _find_loopmidi_exe()
            except Exception as e:
                print(f"[Startup] loopMIDI install error: {e}")

        if exe_path:
            _push(window, 'loopmidi', 'loopMIDI installed', 'done')
        else:
            _push(window, 'loopmidi', 'loopMIDI not found', 'error',
                  'loopMIDI could not be found or installed.\nPlease install it from tobias-erichsen.de')
            return False, None

    # ── Check 2: MidiWarp_OUT port in registry ──
    PORTS_KEY = r"Software\Tobias Erichsen\loopMIDI\Ports"
    _push(window, 'port', 'Checking virtual port', 'active')

    port_existed = False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PORTS_KEY) as k:
            winreg.QueryValueEx(k, OUTPUT_PORT_NAME)
            port_existed = True
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"[Startup] Port registry read error: {e}")

    if port_existed:
        _push(window, 'port', 'Virtual port found', 'done')
    else:
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, PORTS_KEY) as k:
                winreg.SetValueEx(k, OUTPUT_PORT_NAME, 0, winreg.REG_DWORD, 1)
            _push(window, 'port', 'Virtual port created', 'done')
        except Exception as e:
            _push(window, 'port', 'Virtual port failed', 'error',
                  f'Could not create {OUTPUT_PORT_NAME} in registry.\n{e}')
            return False, None

    # ── Check 3: Restart loopMIDI ──
    _push(window, 'loopmidi-restart', 'Restarting loopMIDI', 'active')

    # Set StartMinimized before launching
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Tobias Erichsen\loopMIDI",
                            access=winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, "StartMinimized", 0, winreg.REG_DWORD, 1)
    except Exception as e:
        print(f"[Startup] Could not set StartMinimized: {e}")

    # Kill if running
    try:
        subprocess.call(
            ["taskkill", "/F", "/IM", "loopMIDI.exe"],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass

    # Small gap to let it die cleanly
    time.sleep(0.5)

    # Launch it
    try:
        subprocess.Popen(
            [exe_path],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception as e:
        _push(window, 'loopmidi-restart', 'loopMIDI failed to start', 'error',
              f'Could not launch loopMIDI: {e}')
        return False, None

    # Poll until it appears in the process list (max 8 seconds)
    deadline = time.time() + 8
    running  = False
    while time.time() < deadline:
        if any(p.name().lower() == 'loopmidi.exe'
               for p in psutil.process_iter(['name'])):
            running = True
            break
        time.sleep(0.3)

    if running:
        # Give loopMIDI a moment to load its ports after appearing
        time.sleep(1.0)
        _push(window, 'loopmidi-restart', 'loopMIDI running', 'done')
    else:
        _push(window, 'loopmidi-restart', 'loopMIDI timed out', 'error',
              'loopMIDI did not start in time. Please launch it manually.')
        return False, None

    return True, exe_path


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

def main():
    midi   = MidiHandler()
    engine = ScriptEngine()
    api    = API(midi, engine)

    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(base, "index.html")

    cfg    = load_config()
    bounds = cfg.get("window_bounds", {})
    window = webview.create_window(
        "MidiWarp",
        url=html_path,
        js_api=api,
        width=bounds.get("w", 1100),
        height=bounds.get("h", 700),
        x=bounds.get("x", None),
        y=bounds.get("y", None),
        min_size=(800, 540),
        resizable=True,
        frameless=True,
        easy_drag=False,
    )

    api.set_window(window)

    tray = setup_tray(window, midi, base)

    def on_closed():
        try:
            cfg = load_config()
            cfg["window_bounds"] = {
                "x": window.x, "y": window.y,
                "w": window.width, "h": window.height,
            }
            save_config(cfg)
        except Exception:
            pass
        tray.stop()
        midi.close()

    window.events.closed += on_closed

    def on_loaded():
        def _run():
            ok, exe_path = run_startup_checks(window, base)
            if ok:
                engine.load_scripts()
                midi.open_output()
                try:
                    window.evaluate_js('window.onStartupComplete()')
                except Exception:
                    pass
                t = threading.Thread(target=monitor_loop, args=(midi, engine, api), daemon=True)
                t.start()
                w = threading.Thread(target=port_watchdog, args=(midi, api), daemon=True)
                w.start()
        threading.Thread(target=_run, daemon=True).start()

    window.events.loaded += on_loaded

    webview.start(debug=False)


if __name__ == "__main__":
    main()