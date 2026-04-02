# Vybe

A multi-language compiler and runtime built in Rust. Vybe compiles Visual Basic, JavaScript, Python, C#, Dart, and PHP to a shared bytecode VM with WASM output support. All languages share the same runtime, host functions, and object model — objects created in one language are fully compatible with any other.

## Supported Languages

| Language | Parser | Compiler | Status |
|----------|--------|----------|--------|
| Visual Basic (.vb) | `vybe_parser_basic` | `vybe_compiler_vb` | Mature |
| JavaScript (.js) | `vybe_parser_js` | `vybe_compiler_js` | Mature |
| Python (.py) | `vybe_parser_python` | `vybe_compiler_python` | Mature |
| C# (.cs) | `vybe_parser_csharp` | `vybe_compiler_csharp` | In progress |
| Dart (.dart) | `vybe_parser_dart` | `vybe_compiler_dart` | In progress |
| PHP (.php) | `vybe_parser_php` | `vybe_compiler_php` | In progress |

## Features

- **Multi-language compilation** to a single bytecode format
- **Cross-language interop** — classes, functions, and objects are shared across languages
- **Bytecode VM** with a WASI-compatible host interface
- **WASM output** (`--emit-wasm`) for browser and portable execution
- **Visual Form Designer** with drag-and-drop controls, property editor, and event handlers
- **Project system** — multi-file, multi-language projects via `.vybe` config
- **Sandbox mode** for restricted execution

## Architecture

```
vybe/
├── crates/
│   ├── vybe_bytecode/          # Bytecode VM, opcodes, WASM emitter
│   ├── vybe_compiler_common/   # Shared compiler helpers (classes, functions, I/O)
│   ├── vybe_compiler_vb/       # VB → bytecode
│   ├── vybe_compiler_js/       # JS → bytecode
│   ├── vybe_compiler_python/   # Python → bytecode
│   ├── vybe_compiler_csharp/   # C# → bytecode
│   ├── vybe_compiler_dart/     # Dart → bytecode
│   ├── vybe_compiler_php/      # PHP → bytecode
│   ├── vybe_parser_basic/      # VB parser
│   ├── vybe_parser_js/         # JS parser
│   ├── vybe_parser_python/     # Python parser
│   ├── vybe_parser_csharp/     # C# parser
│   ├── vybe_parser_dart/       # Dart parser
│   ├── vybe_parser_php/        # PHP parser
│   ├── vybe_host/              # Host function registry (I/O, math, strings, etc.)
│   ├── vybe_cli/               # CLI runner (vybec)
│   ├── vybe_ide/               # IDE with Skia renderer
│   ├── vybe_forms/             # Form model and controls
│   ├── vybe_designer/          # Visual form designer
│   ├── vybe_project/           # Project file management
│   └── vybe_widgets/           # UI widget library
└── examples/                   # Example programs per language
```

## Building

```bash
cargo build --release
```

## Usage

```bash
# Compile and run
vybec hello.vb
vybec app.js
vybec script.py
vybec page.php
vybec main.dart

# Dump bytecode
vybec --dump hello.js

# Emit WASM
vybec --emit-wasm hello.js

# Run a multi-language project
vybec project.vybe

# Sandbox mode (restricted host access)
vybec --sandbox untrusted.py
```

## Examples

```js
// JavaScript
function greet(name) {
    console.log("Hello, " + name + "!");
}
greet("World");
```

```python
# Python
def factorial(n):
    if n <= 1:
        return 1
    return n * factorial(n - 1)

print(factorial(6))
```

```php
<?php
$fruits = ["apple", "banana", "cherry"];
foreach ($fruits as $fruit) {
    echo $fruit;
}
```

```vb
' Visual Basic
Sub Main()
    Dim x As Integer = 42
    Console.WriteLine("The answer is " & x)
End Sub
```

## IDE

```bash
cargo build -p vybe_ide --bin skia_ide
```

The IDE includes a code editor with syntax highlighting, a visual form designer, and integrated run/debug support.

## Technology Stack

- **Language**: Rust
- **Bytecode VM**: Custom stack-based VM with WASI-compatible imports
- **WASM**: Component-model output for portable execution
- **GUI Rendering**: tiny-skia
- **Syntax Highlighting**: Monaco Editor (web), tree-sitter (native)

## License

Dual licensed: GPL or Commercial (contact). To be Accepted, Contributions are assumed to be Public Domain
