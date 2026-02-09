# FBEditor — FreeBASIC IDE with GDB Debugger

A production-grade Integrated Development Environment for [FreeBASIC](https://www.freebasic.net/), built with VB.NET and featuring an integrated GDB debugger, AI coding assistant, code outline explorer, and a full-featured Scintilla-powered editor.

![.NET Framework 4.8](https://img.shields.io/badge/.NET_Framework-4.8-purple)
![Platform](https://img.shields.io/badge/Platform-Windows-blue)
![Version](https://img.shields.io/badge/Version-4.0.0-green)
![License](https://img.shields.io/badge/License-Permissive-orange)

---

## Features

### Editor

- **Scintilla-powered editor** (ScintillaNET 3.6.3) with full FreeBASIC syntax highlighting — keywords, preprocessor directives, built-in functions, strings, numbers, and comments
- **Auto-complete** with FreeBASIC keywords, built-in functions, and all identifiers from your source code (types, constants, variables, sub/function names, enum values, `#define` names, and more)
- **Multi-file tabbed editing** — open and work on multiple `.bas` and `.bi` files simultaneously, switch with `Ctrl+Tab` / `Ctrl+Shift+Tab`
- **Code outline panel** — live-updating tree view showing the structure of your code, organized into categories:
  - **Types** — `TYPE ... END TYPE` blocks
  - **Enums** — `ENUM ... END ENUM` blocks
  - **Constants** — `CONST` declarations and `#DEFINE` macros
  - **Procedures** — `SUB` and `FUNCTION` definitions with signatures and return types
  - **Global Variables** — `DIM SHARED`, `COMMON SHARED`, and module-level variables
  - **Global Arrays** — `DIM SHARED` arrays and `REDIM SHARED` dynamic arrays
  - **Variables / Arrays** — local variables within procedures
  - **Declares** — `DECLARE SUB` / `DECLARE FUNCTION` forward declarations
  - **Properties** — `PROPERTY` definitions
  - Double-click any item to jump directly to that line
- **Find & Replace** with match case, whole word, regex, and wrap-around options
- **Go To Line** (`Ctrl+G`)
- **Bookmarks** — toggle, navigate next/previous, clear all
- **Comment / Uncomment** blocks
- **Encoding support** — ANSI (Windows-1252), UTF-8, and UTF-8 with BOM, with configurable default
- **Line ending display** in status bar (CRLF/LF)
- **Zoom** in/out/reset
- **Customizable editor font**
- **Toggle options** — line numbers, word wrap, whitespace visibility, indentation guides
- **Recent files** menu
- **Drag & drop** — drop `.bas` / `.bi` files onto the editor to open them

### Build System

- **Compile** (`Ctrl+F5`), **Compile & Run** (`F6`), **Run without compile** (`Ctrl+F6`)
- **Quick Run** — compile and run in one step
- **Syntax Check Only** — validate code without producing an executable
- **Build Options dialog** with full control over the FreeBASIC compiler:
  - **Target type** — Executable, DLL, Static Library
  - **Optimization level** — O0 through O3
  - **Error checking** — none, standard, pedantic
  - **Dialect** — fb (default), fblite, qb, deprecated
  - **Code generation** — GAS, GCC, LLVM
  - **Warnings** — off, all, pedantic, errors
  - **Architecture** — 32-bit / 64-bit
  - **FPU mode** — SSE, x87, NEON
  - **Debug info**, verbose output, show commands, generate map file, emit ASM, keep intermediate files
  - **Custom compiler flags**, library paths, and include paths
  - **Separate 32-bit and 64-bit FBC paths**
- **Live command preview** — see the exact `fbc` command line as you change options
- **Clickable error output** — click compiler errors to jump to the relevant source line

### Integrated GDB Debugger

A fully integrated GDB front-end designed specifically for FreeBASIC programs:

- **Start / Continue** (`F5`), **Stop** (`Shift+F5`), **Pause**
- **Step Over** (`F10`), **Step Into** (`F11`), **Step Out** (`Shift+F11`)
- **Run to Cursor** (`Ctrl+F10`)
- **Breakpoints** — toggle (`F9`), clear all, persistent across sessions, shown as red markers in the editor margin
- **Current line highlighting** — yellow marker shows the current execution point
- **Locals / Watch panel** — automatically displays local variables and function arguments at each break
- **Watch expressions** — add custom expressions to monitor during debugging
- **Call stack** — view the full call stack with file and line information
- **Debug output** — separate panel for GDB/inferior output
- **GDB command line** — direct GDB/MI command input for advanced users
- **Debug toolbar** — quick-access buttons for Debug, Stop, Pause, Step Over, Step Into, Step Out, and Breakpoint toggle

### AI Chat (Claude Integration)

Built-in AI coding assistant powered by Claude (Anthropic API):

- **Ask questions** about FreeBASIC programming directly within the IDE
- **Send code** — include your current source file for context-aware responses
- **Insert Code** — paste AI-suggested code at the cursor position
- **Copy Reply** — copy the AI response to clipboard
- **Replace All** — replace the entire editor content with AI-generated code
- **New File** — create a new file from AI-generated code
- **Clear Chat** — reset conversation history
- **Quick send** — `Ctrl+Enter` in the input box
- Requires an Anthropic API key stored in `api_key.txt` in the application directory

### Themes

- **Dark theme** and **Light theme** — toggle with a single menu click
- Themes apply to the entire IDE: editor, panels, menus, status bar, and dialogs
- Scintilla syntax colors are optimized for each theme

### Other

- **Project Explorer** — tree view of open files
- **Collapsible panels** — show/hide project explorer and output panel independently
- **Status bar** — current line/column position, line count, insert/overwrite mode, encoding, line ending style, debug state
- **Persistent settings** — window size/position, editor preferences, compiler paths, recent files, theme, and more are saved between sessions in JSON format
- **FreeBASIC Help** — quick access to local FB help file (`F1`) and online documentation

---

## Screenshots

*Coming soon*

---

## Requirements

- **Windows 7 SP1** or later (32-bit or 64-bit)
- **.NET Framework 4.8** — [Download](https://dotnet.microsoft.com/download/dotnet-framework/net48)
- **FreeBASIC Compiler** — [Download](https://www.freebasic.net/wiki/CompilerInstalling)
- **GDB** (for debugging) — included with FreeBASIC for Windows, or install separately via [MinGW](https://www.mingw-w64.org/)
- **Anthropic API Key** (optional, for AI Chat) — [Get one here](https://console.anthropic.com/)

---

## Installation

### From Installer

1. Download the latest `FBEditor_Setup_4.0.0.exe` from [Releases](https://github.com/ronen-blumberg/FBEditor/releases/tag/v4.0.0)
2. Run the installer — it will check for .NET Framework 4.8 and guide you through setup
3. FBEditor installs to `C:\Program Files (x86)\FBEditor` by default
4. Optional: associate `.bas` and `.bi` files with FBEditor during installation

### From Source

1. Clone the repository:
   ```
   git clone https://github.com/ronen-blumberg/FBEditor.git
   ```
2. Open `FBEditor.sln` in Visual Studio 2022/2026
3. Ensure the following NuGet packages are installed:
   - `jacobslusser.ScintillaNET` (3.6.3)
   - `Newtonsoft.Json` (13.0.3)
4. Build in Release mode (target: .NET Framework 4.8)
5. The output will be in `bin\Release\net48\`

### Building the Installer

1. Install [Inno Setup](https://jrsoftware.org/isinfo.php) (6.x or later)
2. Open `Installer\FBEditor_Setup.iss`
3. Update the `MyAppSource` path at the top to point to your `bin\Release\net48` folder
4. Compile with the Inno Setup Compiler
5. The installer will be output to the `Output\` directory

---

## Getting Started

1. **Set the FreeBASIC compiler path** — go to `Tools → FreeBASIC Compiler Path...` and browse to your `fbc.exe`
2. **Create a new file** (`Ctrl+N`) or open an existing `.bas` file (`Ctrl+O`)
3. **Write your code** — enjoy syntax highlighting, auto-complete (`Ctrl+Space`), and the code outline
4. **Compile & Run** — press `F6` to compile and run your program
5. **Debug** — press `F9` to set a breakpoint, then `F5` to start debugging

### Setting Up AI Chat

1. Get an API key from [Anthropic Console](https://console.anthropic.com/)
2. Create a file named `api_key.txt` in the FBEditor application directory
3. Paste your API key into the file (single line, no spaces)
4. Open the AI Chat panel from `View → AI Chat Panel`

### Setting Up GDB

1. Go to `Debug → Set GDB Path...`
2. Browse to your `gdb.exe` — typically located in the FreeBASIC `bin\win32` or `bin\win64` directory
3. Make sure your Build Options have **Debug Info** (`-g`) enabled

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+N` | New File |
| `Ctrl+O` | Open File |
| `Ctrl+S` | Save |
| `Ctrl+W` | Close File |
| `Ctrl+Tab` | Next Tab |
| `Ctrl+Shift+Tab` | Previous Tab |
| `Ctrl+Z` / `Ctrl+Y` | Undo / Redo |
| `Ctrl+Space` | Auto-Complete |
| `Ctrl+F` | Find |
| `Ctrl+H` | Find & Replace |
| `F3` | Find Next |
| `Ctrl+G` | Go To Line |
| `F4` | Refresh Code Outline |
| `Ctrl+F5` | Compile |
| `F6` | Compile & Run |
| `Ctrl+F6` | Run (No Compile) |
| `F5` | Start / Continue Debugging |
| `Shift+F5` | Stop Debugging |
| `F9` | Toggle Breakpoint |
| `F10` | Step Over |
| `F11` | Step Into |
| `Shift+F11` | Step Out |
| `Ctrl+F10` | Run to Cursor |
| `F1` | FreeBASIC Help |

---

## Project Structure

```
FBEditor/
├── FBEditor.sln                    Solution file
├── FBEditor.vbproj                 Project file (.NET Framework 4.8)
├── Program.vb                      Entry point
├── Forms/
│   ├── MainForm.vb                 Main IDE window (1,785 lines)
│   ├── FindReplaceForm.vb          Find & Replace dialog
│   ├── GoToLineForm.vb             Go To Line dialog
│   ├── BuildOptionsForm.vb         Compiler options dialog
│   └── AboutForm.vb                About dialog
├── Modules/
│   ├── GDBDebugger.vb              GDB/MI protocol integration (943 lines)
│   ├── BuildSystem.vb              Compiler invocation & output parsing
│   ├── CodeOutline.vb              Source code parser for outline tree
│   ├── SyntaxConfig.vb             FreeBASIC keywords & syntax definitions
│   ├── ThemeManager.vb             Dark/Light theme engine
│   ├── AIChatManager.vb            Claude API integration
│   ├── AppSettings.vb              Application-wide settings
│   └── UserSettings.vb             Per-user persistent settings (JSON)
├── Resources/
│   ├── FBEditor.ico                Application icon
│   └── LICENSE.txt                 License file
└── Installer/
    └── FBEditor_Setup.iss          Inno Setup installer script
```

**Total: ~5,900 lines of VB.NET**

---

## Dependencies

| Library | Version | License | Purpose |
|---|---|---|---|
| [ScintillaNET](https://github.com/jacobslusser/ScintillaNET) | 3.6.3 | MIT | Code editor component |
| [Newtonsoft.Json](https://www.newtonsoft.com/json) | 13.0.3 | MIT | Settings serialization |
| [Scintilla](https://www.scintilla.org/) | (bundled) | Scintilla License | Native text editing engine |

---

## License

Copyright © 2026 Ronen Blumberg. All rights reserved.

Permission is granted to use, copy, and distribute this software for any purpose, provided that the copyright notice and license appear in all copies or substantial portions of the software.

This software is provided "as-is" without warranty of any kind.

See [LICENSE](Resources/LICENSE.txt) for full details.

---

## Author

**Ronen Blumberg**

Built with ❤️ for the FreeBASIC community.
