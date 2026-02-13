# Vybe Compiler/Interpreter with Visual Form Editor

A Visual Basic-style development environment built from scratch in Rust, featuring a parser, interpreter, and visual form designer.

## Features

- **Vybe Basic Parser**: Full support for Basic syntax including:
  - Variables, assignments, and expressions
  - Control flow (If/Then/Else, For/Next, While/Wend, Do/Loop)
  - Procedures (Sub/Function) with parameters
  - Classes with properties and methods
  - Event handlers

- **Interpreter**: Tree-walking interpreter with:
  - Expression evaluation
  - Variable scoping
  - Function calls
  - Built-in functions (MsgBox, Len, Left, Right, Mid, UCase, LCase)
  - Event system for control events

- **Visual Form Designer**:
  - Drag-and-drop control placement
  - Property editor
  - Toolbox with standard controls (Button, Label, TextBox, CheckBox, RadioButton)
  - Grid-based designer canvas
  - Code editor view

- **Project System**:
  - JSON-based project files
  - Form serialization
  - Multiple forms and modules support

## Architecture

```
vb/
├── crates/
│   ├── vybe_parser/      # PEG parser using pest
│   ├── vybe_runtime/     # Tree-walking interpreter
│   ├── vybe_forms/       # Form model and controls
│   ├── vybe_editor/      # Visual editor (iced GUI)
│   └── vybe_project/     # Project file management
└── examples/            # Sample vybe programs
```

## Building

```bash
cd vb
cargo build --release
```

## Running the Editor

```bash
cargo run --bin vybe_editor
```

## Running Example Code

```bash
# Parse and execute a IRIS file
cargo run --example parse_and_run examples/hello_world.vb

```

## Usage

1. **Create a New Project**: Click "New" to create a new vybe Basic project
2. **Add Controls**: Click a control type from the toolbox, then click on the form to place it
3. **Edit Properties**: Select a control to view and edit its properties
4. **Write Code**: Click "View Code" to write event handlers
5. **Run**: Click "Run" to execute your program

## Example vybe Code

```vb
' Button click event handler
Sub btnHello_Click()
    MsgBox("Hello, World!")
End Sub

' Function with parameters
Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
```

## Technology Stack

- **Language**: Rust 2021
- **Parser**: pest (PEG parser)
- **GUI**: Dioxus (retained-mode GUI framework)
- **Serialization**: serde + serde_json
- **Syntax Highlight**: Monaco Editor

## Roadmap

- [x] Basic parser for VB syntax
- [x] Tree-walking interpreter
- [x] Form model with controls
- [x] Visual designer with iced
- [x] Property editor
- [x] Event system
- [x] More controls (ComboBox, ListBox, Frame, PictureBox)
- [ ] File dialogs for Open/Save
- [ ] Debugger
- [ ] Bytecode VM for better performance
- [x] IntelliSense/autocomplete
- [ ] WASM compilation

## License

Dual licensed: GPL or Commercial (contact). To be Accepted, Contributions are assumed to be Public Domain
