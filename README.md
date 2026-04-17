# MonoStereoToggle

A lightweight Windows 11 utility to instantly switch your audio output between **Mono** and **Stereo** modes — the same setting found in Windows Settings → Accessibility → Audio.

![License](https://img.shields.io/github/license/samaBR85/MonoStereoToggle)
![Release](https://img.shields.io/github/v/release/samaBR85/MonoStereoToggle)
![Platform](https://img.shields.io/badge/platform-Windows%2011-blue)
![Language](https://img.shields.io/badge/language-C%23-green)

---

## Features

- 🔊 One-click toggle between MONO and STEREO
- ⌨️ Global keyboard shortcut (customizable)
- 🖥️ Visual overlay notification on toggle
- 🔔 System tray icon showing current state
- 📋 Lists all available audio output devices
- 🚀 Auto-start with Windows (no UAC prompt on login)
- 🌙 Dark UI

---

## Download

👉 **[Latest Release](https://github.com/samaBR85/MonoStereoToggle/releases/latest)**

Download `MonoStereoToggle.exe` — no installation required.

---

## Requirements

- Windows 11
- [.NET 8 Runtime (x64)](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)

---

## Usage

1. Run `MonoStereoToggle.exe` — administrator privileges are requested once on launch
2. Click the button to toggle between **STEREO** and **MONO**
3. Optionally set a global keyboard shortcut and enable auto-start with Windows

---

## Building from Source

```bash
git clone https://github.com/samaBR85/MonoStereoToggle.git
cd MonoStereoToggle
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o publish
```

The executable will be at `publish\MonoStereoToggle.exe`.

---

## License

[MIT](LICENSE) © 2026 [samaBR](https://github.com/samaBR85)
