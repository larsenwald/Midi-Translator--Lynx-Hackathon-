# MidiWarp

Real-time MIDI transformation for musicians who want more control.

MidiWarp sits between your MIDI keyboard and your DAW. It intercepts your signal, runs it through scripts you write, and sends the result to a virtual MIDI port your DAW listens to. Transpose notes, remap controls, filter velocities, convert pitch bend to CC, whatever you need, without touching your DAW settings.

---

## How it works

Scripts are plain Python. No special syntax. If you can write a basic `if` statement, you can write a script.

---

## Features

- **Live event feed:** see every MIDI event in real time as you play, including what changed after your scripts ran
- **Script pipeline:** write multiple scripts that run in sequence; drag to reorder
- **Simple scripting API:** read and mutate events with human-readable variables (`note`, `velocity`, `cc_num`, `pitch`), no raw bytes or hex
- **Note name constants:** use `C4`, `Fs3`, `Bb5` directly instead of numbers
- **AI script generation:** describe what you want in plain English and get a working script back (requires a free Gemini API key)
- **Built-in editor:** line numbers, syntax hints, and a full scripting reference
- **Auto setup:** installs and configures loopMIDI automatically on first launch
- **Runs in the background:** minimizes to system tray

---

## Installation

1. Download the latest `MidiWarpInstaller.exe` from [Releases](../../releases)
2. Run the installer
3. MidiWarp will handle installing loopMIDI automatically on first launch

No manual driver setup needed.

---

## Quick start

1. Launch MidiWarp
2. Click the orb on the home screen and pick your MIDI keyboard from the dropdown
3. Play some notes and watch events show up in the feed
4. Go to the **Scripts** tab to start writing transforms

---

## Example scripts

**Transpose all notes up an octave:**
```python
if event_type in ('note_on', 'note_off'):
    note += 12
```

**Remap C5, E5, F5 to D5:**
```python
if event_type in ('note_on', 'note_off') and note in (C5, E5, F5):
    note = D5
```

**Convert pitch bend to CC11 (Expression), clamped at zero:**
```python
if event_type == 'pitchwheel':
    send_cc(11, max(0, int(pitch / 8191 * 127)))
    block()
```

**Drop ghost notes (velocity below 20):**
```python
if event_type == 'note_on' and velocity < 20:
    block()
```

---

## Tech stack

- [Python](https://python.org)
- [pywebview](https://pywebview.flowrl.com)
- [mido](https://mido.readthedocs.io) + [python-rtmidi](https://spotlightkid.github.io/python-rtmidi/)
- [loopMIDI](https://www.tobias-erichsen.de/software/loopmidi.html) by Tobias Erichsen
- [Google Gemini API](https://ai.google.dev)
- [psutil](https://psutil.readthedocs.io)
- [Inno Setup](https://jrsoftware.org/isinfo.php)

---

## Contributors

Built for **Lynx Hack 2026** at the University of Colorado Denver by **Team YAMI**.

---

## License

MIT. Do whatever you want with it.
