# MidiWarp

**Real-time MIDI transformation for musicians who want more control.**

MidiWarp lets you intercept your MIDI keyboard's signal and reshape it before it reaches your DAW -- transpose notes, remap controls, filter velocities, convert pitch bend to CC, and more. All without touching your DAW's settings.

If you've ever wished you could just *make* your keyboard do something slightly different without rewiring your whole project, MidiWarp is for you.

---

## How it works

MidiWarp sits between your MIDI keyboard and your DAW. It receives your keyboard's signal, runs it through any scripts you've written, and sends the transformed result to a virtual MIDI port that your DAW listens to.

Scripts are plain Python -- no special syntax to learn. If you can write a basic `if` statement, you can write a script.

---

## Features

- **Live event feed** -- see every MIDI event in real time as you play, including what changed after your scripts ran
- **Script pipeline** -- write multiple scripts that run in sequence on every event; drag to reorder them
- **Clean scripting API** -- read and mutate events using human-readable variables (`note`, `velocity`, `cc_num`, `pitch`) with no raw bytes or hex
- **Note name constants** -- use `C4`, `Fs3`, `Bb5` directly in scripts instead of numbers
- **AI script generation** -- describe what you want in plain English and get a working script back (requires a free Gemini API key from Google AI Studio)
- **Script editor** -- built-in code editor with line numbers, syntax hints, and a full scripting reference
- **Auto setup** -- installs and configures loopMIDI automatically on first launch
- **Runs in the background** -- minimizes to system tray, stays out of your way

---

## Installation

1. Download the latest `MidiWarpInstaller.exe` from [Releases](../../releases)
2. Run the installer -- it will walk you through setup
3. MidiWarp will handle installing loopMIDI (the virtual MIDI driver it depends on) automatically on first launch

That's it. No manual driver setup required.

---

## Quick start

1. Launch MidiWarp
2. Click the orb on the home screen and select your MIDI keyboard from the dropdown
3. Play some notes -- you should see events appear in the feed
4. Head to the **Scripts** tab to start writing transforms

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
- [pywebview](https://pywebview.flowrl.com) -- native window with a web UI
- [mido](https://mido.readthedocs.io) + [python-rtmidi](https://spotlightkid.github.io/python-rtmidi/) -- MIDI I/O
- [loopMIDI](https://www.tobias-erichsen.de/software/loopmidi.html) -- virtual MIDI port (by Tobias Erichsen)
- [Google Gemini API](https://ai.google.dev) -- AI script generation
- [psutil](https://psutil.readthedocs.io) -- process and service management
- [Inno Setup](https://jrsoftware.org/isinfo.php) -- Windows installer

---

## Contributors

MidiWarp was built for **Lynx Hack 2026** at the University of Colorado Denver by **Team YAMI**.

---

## License

MIT License -- do whatever you want with it.
