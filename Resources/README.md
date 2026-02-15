# FBEditor v5.0.0 — FreeBASIC IDE with Visual Form Designer

A production-grade Integrated Development Environment for [FreeBASIC](https://www.freebasic.net/), built with VB.NET and featuring a **Window9 Visual Form Designer**, integrated GDB debugger, AI coding assistant, code outline explorer, and a full-featured Scintilla-powered editor.

![.NET Framework 4.8](https://img.shields.io/badge/.NET_Framework-4.8-purple)
![Platform](https://img.shields.io/badge/Platform-Windows-blue)
![Version](https://img.shields.io/badge/Version-5.0.0-green)
![License](https://img.shields.io/badge/License-Permissive-orange)

---

## What's New in v5.0.0

### Window9 Visual Form Designer (NEW)

FBEditor v5.0.0 introduces a full WYSIWYG visual form designer for the [Window9](https://users.freebasic-portal.de/tjf/Projekte/Window9/doc/Window9.html) GUI library — the first of its kind for FreeBASIC. Design your GUI visually, then generate complete FreeBASIC + Window9 source code with a single click.

**Designer Features:**
- Drag-and-drop placement of 23 gadget types on a visual canvas
- Grid snapping with configurable grid size
- Resize handles on all 8 edges/corners
- Multi-select with Ctrl+Click or rubber-band selection
- Arrow key nudging (snaps to grid)
- Full Undo/Redo (Ctrl+Z / Ctrl+Y)
- Copy/Cut/Paste/Duplicate gadgets (Ctrl+C/X/V/D)
- Right-click context menu with all editing operations
- Lock gadgets to prevent accidental movement
- Bring to Front / Send to Back (Z-order control)
- Alignment tools: Left, Right, Top, Bottom, Center Horizontally/Vertically
- Make Same Size: Width, Height, or Both
- Space Evenly: Horizontal and Vertical distribution
- Tab Order display overlay showing gadget creation sequence
- Select All (Ctrl+A)
- Save/Load form designs as `.w9form` JSON files
- Preview generated code before saving

**Supported Gadget Types (23):**
- **Button** — Standard push button with click events
- **TextLabel** — Static text display
- **Editor** — Multi-line text editor with optional scrollbars, word wrap, and read-only mode
- **StringInput** — Single-line text input with optional password masking
- **CheckBox** — Toggle checkbox with checked/unchecked state
- **OptionButton** — Radio button for mutually exclusive choices
- **ComboBox** — Drop-down selection list with initial items
- **ListBox** — Scrollable selection list with initial items
- **GroupBox** — Visual container for organizing related controls
- **ImageBox** — Image display area
- **ProgressBar** — Progress indicator with min/max/value
- **ScrollBar** — Horizontal or vertical scroll bar
- **TrackBar** — Slider control with min/max/value
- **SpinBox** — Numeric up/down spinner
- **TreeView** — Hierarchical tree display with optional lines, buttons, and checkboxes
- **ListView** — Multi-column list display with optional checkboxes and full-row select
- **StatusBar** — Window status bar with configurable fields
- **PanelTab** — Tabbed panel container with named tab pages
- **Container** — Generic panel/container
- **Splitter** — Resizable divider between controls
- **Calendar** — Date picker calendar control
- **HyperLink** — Clickable URL link
- **WebBrowser** — Embedded web browser control

**Property Panel — 30+ Properties per Gadget:**
- **Layout:** EnumName, X, Y, Width, Height
- **Content:** Text, Items (for ComboBox/ListBox), TabNames (for PanelTab), ImagePath
- **Appearance:** FontName (dropdown with all system fonts), FontSize, BackColor, ForeColor, Tooltip
- **Behavior:** IsReadOnly, WordWrap, IsChecked, IsPassword, IsEnabled, IsVisible, Min/Max/Value, Orientation
- **Editor:** HasVScroll, HasHScroll
- **TreeView/ListView:** HasLines, HasButtons, HasCheckBoxes, FullRowSelect
- **Events:** OnClickEvent, OnChangeEvent, OnDoubleClickEvent
- **Advanced:** Style, ExStyle, Tag

**Menu Designer:**
- Visual menu bar editor for creating application menus
- Add, remove, and reorder menu items
- Assign keyboard shortcut accelerators
- Add separators between menu groups
- Generated code includes full menu creation and event dispatch

**Code Generator — Complete Application Output:**
- Generates complete, compilable FreeBASIC + Window9 source code
- Proportional resize system with ScaleX/ScaleY macros
- Linux compatibility defines for cross-platform builds
- Gadget enum IDs with customizable start values
- Proper Window9 API calls for every gadget type and property
- SetGadgetFont with explicit font name and size for every gadget
- Event handler stubs with helpful TODO comments and API usage examples
- Smart event dispatch in the main event loop
- Window resize handler with proportional gadget scaling
- Context-aware Window9 API Quick Reference appended as comments (only shows APIs for gadgets you actually used)

**Generated Event Handlers Include:**
- Button click handlers
- CheckBox/OptionButton toggle handlers with GetGadgetState
- ComboBox/ListBox selection change handlers with index and text retrieval
- Editor/StringInput text change handlers
- ScrollBar/TrackBar/SpinBox value change handlers
- ListView/TreeView double-click handlers
- Window resize and close events
- Menu item click handlers
- Timer events

---

## Features

### Editor

- **Scintilla-powered editor** (ScintillaNET 3.6.3) with full FreeBASIC syntax highlighting — keywords, preprocessor directives, built-in functions, strings, numbers, and comments
- **Auto-complete** with FreeBASIC keywords, built-in functions, and all identifiers from your source code (types, constants, variables, sub/function names, enum values, `#define` names, and more)
- **Multi-file tabbed editing** — open and work on multiple `.bas` and `.bi` files simultaneously, switch with `Ctrl+Tab` / `Ctrl+Shift+Tab`
- **Code outline panel** — live-updating tree view showing the structure of your code, organized into categories:
  - Types, Enums, Constants, Procedures, Global Variables, Global Arrays, Variables/Arrays, Declares, Properties
  - Double-click any item to jump directly to that line
- **Code folding** — collapse/expand Sub/Function, Type, Enum, #If blocks, and multi-line comments
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
- **Project types** — Console Application, GUI Application, Window9 Forms Application
- **Build Options dialog** with full control over the FreeBASIC compiler:
  - Target type — Executable, DLL, Static Library
  - Optimization level — O0 through O3
  - Error checking — none, standard, pedantic
  - Dialect — fb (default), fblite, qb, deprecated
  - Code generation — GAS, GCC, LLVM
  - Warnings — off, all, pedantic, errors
  - Architecture — 32-bit / 64-bit
  - FPU mode — SSE, x87, NEON
  - Debug info, verbose output, show commands, generate map file, emit ASM, keep intermediate files
  - Custom compiler flags, library paths, and include paths
  - Separate 32-bit and 64-bit FBC paths
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
- **Debug toolbar** — quick-access buttons for all debug operations

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
- Themes apply to the entire IDE: editor, panels, menus, status bar, dialogs, and form designer
- Scintilla syntax colors are optimized for each theme

### Other

- **Project Explorer** — tree view of open files
- **Collapsible panels** — show/hide project explorer and output panel independently
- **Status bar** — current line/column position, line count, insert/overwrite mode, encoding, line ending style, debug state
- **Persistent settings** — window size/position, editor preferences, compiler paths, recent files, theme, and more are saved between sessions in JSON format
- **FreeBASIC Help** — quick access to local FB help file (`F1`) and online documentation

---

## Requirements

- **Windows 7 SP1** or later (32-bit or 64-bit)
- **.NET Framework 4.8** — [Download](https://dotnet.microsoft.com/download/dotnet-framework/net48)
- **FreeBASIC Compiler** — [Download](https://www.freebasic.net/wiki/CompilerInstalling)
- **Window9 Library** (for Visual Form Designer output) — [Download](https://users.freebasic-portal.de/tjf/Projekte/Window9/doc/Window9.html)
- **GDB** (for debugging) — included with FreeBASIC for Windows, or install separately via [MinGW](https://www.mingw-w64.org/)
- **Anthropic API Key** (optional, for AI Chat) — [Get one here](https://console.anthropic.com/)

---

## Installation

### From Installer

1. Download the latest `FBEditor_v5.0_Pro_Setup.exe` from [Releases](https://github.com/ronen-blumberg/FBEditor/releases)
2. Run the installer — it will check for .NET Framework 4.8 and guide you through setup
3. FBEditor installs to `C:\Program Files (x86)\FBEditor` by default
4. Optional: associate `.bas`, `.bi`, and `.w9form` files with FBEditor during installation

### From Source

1. Clone the repository:
   ```
   git clone https://github.com/ronen-blumberg/FBEditor.git
   ```
2. Open `FBEditor.sln` in Visual Studio 2022 or later
3. Ensure the following NuGet packages are installed:
   - `jacobslusser.ScintillaNET` (3.6.3)
   - `Newtonsoft.Json` (13.0.3)
4. Build in Release mode (target: x86, .NET Framework 4.8)
5. The output will be in `bin\x86\Release\net48\`

### Building the Installer

1. Install [Inno Setup](https://jrsoftware.org/isinfo.php) (6.x or later)
2. Open `Installer\FBEditor_Setup.iss`
3. Update the `BuildOutput` path at the top to point to your build output folder
4. Compile with the Inno Setup Compiler
5. The installer will be created in the `Installer\` directory

---

## Getting Started

### Basic Usage

1. **Set the FreeBASIC compiler path** — go to `Tools → FreeBASIC Compiler Path...` and browse to your `fbc.exe`
2. **Create a new file** (`Ctrl+N`) or open an existing `.bas` file (`Ctrl+O`)
3. **Write your code** — enjoy syntax highlighting, auto-complete (`Ctrl+Space`), and the code outline
4. **Compile & Run** — press `F6` to compile and run your program
5. **Debug** — press `F9` to set a breakpoint, then `F5` to start debugging

### Using the Visual Form Designer

1. **Open the Form Designer** — go to `Tools → Window9 Form Designer` or click the Form Designer button
2. **Select a gadget** from the Toolbox panel on the left (Button, Editor, ComboBox, etc.)
3. **Draw on the canvas** — click and drag to place the gadget at your desired size and position
4. **Configure properties** — use the Property Panel on the right to set text, font, colors, events, and behavior
5. **Design menus** — click "Menus..." in the toolbar to open the Menu Designer
6. **Set events** — check OnClickEvent, OnChangeEvent, or OnDoubleClickEvent in the Events category to generate handler stubs
7. **Generate code** — click "Generate Code" to produce a complete FreeBASIC + Window9 source file
8. **Save your design** — click "Save Design" to save as a `.w9form` file for later editing
9. **Compile** — the generated code compiles directly with `fbc -s gui yourfile.bas` (requires Window9 library)

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

### General

| Shortcut | Action |
|---|---|
| `Ctrl+N` | New File |
| `Ctrl+O` | Open File |
| `Ctrl+S` | Save |
| `Ctrl+W` | Close File |
| `Ctrl+Tab` | Next Tab |
| `Ctrl+Shift+Tab` | Previous Tab |

### Editor

| Shortcut | Action |
|---|---|
| `Ctrl+Z` / `Ctrl+Y` | Undo / Redo |
| `Ctrl+Space` | Auto-Complete |
| `Ctrl+F` | Find |
| `Ctrl+H` | Find & Replace |
| `F3` | Find Next |
| `Ctrl+G` | Go To Line |
| `F4` | Refresh Code Outline |

### Build & Run

| Shortcut | Action |
|---|---|
| `Ctrl+F5` | Compile |
| `F6` | Compile & Run |
| `Ctrl+F6` | Run (No Compile) |
| `F1` | FreeBASIC Help |

### Debugger

| Shortcut | Action |
|---|---|
| `F5` | Start / Continue Debugging |
| `Shift+F5` | Stop Debugging |
| `F9` | Toggle Breakpoint |
| `F10` | Step Over |
| `F11` | Step Into |
| `Shift+F11` | Step Out |
| `Ctrl+F10` | Run to Cursor |

### Form Designer

| Shortcut | Action |
|---|---|
| `Ctrl+C` | Copy selected gadget |
| `Ctrl+X` | Cut selected gadget |
| `Ctrl+V` | Paste gadget |
| `Ctrl+D` | Duplicate selected gadget |
| `Ctrl+A` | Select all gadgets |
| `Ctrl+Z` / `Ctrl+Y` | Undo / Redo |
| `Delete` | Delete selected gadget(s) |
| `Arrow Keys` | Nudge gadget position (snaps to grid) |
| `Escape` | Deselect / Cancel placement |

---

## Project Structure

```
FBEditor/
├── FBEditor.sln                        Solution file
├── FBEditor.vbproj                     Project file (.NET Framework 4.8, x86)
├── Program.vb                          Entry point
├── Forms/
│   ├── MainForm.vb                     Main IDE window (2,066 lines)
│   ├── FormDesignerPanel.vb            Form designer host panel with toolbar (542 lines)
│   ├── W9DesignerCanvas.vb             Visual design surface / canvas (1,321 lines)
│   ├── W9ToolboxPanel.vb               Gadget toolbox with 23 types (195 lines)
│   ├── W9PropertyPanel.vb              Property grid with 30+ properties (855 lines)
│   ├── W9MenuDesigner.vb               Menu bar visual editor (277 lines)
│   ├── FindReplaceForm.vb              Find & Replace dialog (277 lines)
│   ├── GoToLineForm.vb                 Go To Line dialog (69 lines)
│   ├── BuildOptionsForm.vb             Compiler options dialog (554 lines)
│   └── AboutForm.vb                    About dialog (139 lines)
├── Modules/
│   ├── W9CodeGenerator.vb              FreeBASIC + Window9 code generator (881 lines)
│   ├── W9GadgetInfo.vb                 Gadget type definitions & registry (423 lines)
│   ├── GDBDebugger.vb                  GDB/MI protocol integration (943 lines)
│   ├── BuildSystem.vb                  Compiler invocation & output parsing (262 lines)
│   ├── CodeOutline.vb                  Source code parser for outline tree (434 lines)
│   ├── FoldingManager.vb               Code folding engine (411 lines)
│   ├── SyntaxConfig.vb                 FreeBASIC keywords & syntax definitions (166 lines)
│   ├── ThemeManager.vb                 Dark/Light theme engine (358 lines)
│   ├── AIChatManager.vb                Claude API integration (287 lines)
│   ├── ProjectManager.vb               Project file management (166 lines)
│   ├── AppSettings.vb                  Application-wide settings (380 lines)
│   └── UserSettings.vb                 Per-user persistent settings (488 lines)
├── Resources/
│   ├── FBEditor.ico                    Application icon
│   ├── LICENSE.txt                     License file
│   └── README.md                       This file
└── Installer/
    └── FBEditor_Setup.iss              Inno Setup installer script
```

**Total: ~11,500 lines of VB.NET**

---

## Dependencies

| Library | Version | License | Purpose |
|---|---|---|---|
| [ScintillaNET](https://github.com/jacobslusser/ScintillaNET) | 3.6.3 | MIT | Code editor component |
| [Newtonsoft.Json](https://www.newtonsoft.com/json) | 13.0.3 | MIT | Settings & form design serialization |
| [Scintilla](https://www.scintilla.org/) | (bundled) | Scintilla License | Native text editing engine |

---

## Version History

### v5.0.0 (February 2026)
- **NEW:** Window9 Visual Form Designer with 23 gadget types
- **NEW:** Property panel with 30+ configurable properties per gadget
- **NEW:** Menu Designer for creating application menus
- **NEW:** FreeBASIC + Window9 code generator with proportional resize support
- **NEW:** Event system with OnClick, OnChange, OnDoubleClick for all gadget types
- **NEW:** Designer UX: context menu, alignment tools, spacing tools, z-order, lock, tab order display
- **NEW:** Font name dropdown picker with all system fonts
- **NEW:** `.w9form` save/load format for form designs
- **NEW:** Window9 API Quick Reference in generated code

### v4.3.0 (February 2026)
- Code folding support
- Enhanced build options

### v4.0.0 (February 2026)
- Initial public release
- Scintilla editor with FreeBASIC syntax highlighting
- Integrated GDB debugger
- AI Chat with Claude integration
- Code outline explorer
- Dark/Light themes

---

## License

Copyright © 2026 Ronen Blumberg. All rights reserved.

Permission is granted to use, copy, and distribute this software for any purpose, provided that the copyright notice and license appear in all copies or substantial portions of the software.

This software is provided "as-is" without warranty of any kind.

See [LICENSE](LICENSE.txt) for full details.

---

## Author

**Ronen Blumberg**

Built with ❤️ for the FreeBASIC community.
