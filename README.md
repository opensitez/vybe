# Vybe

Vybe is a Rust workspace for compiling multiple source languages into a shared bytecode VM, running them through a common host/runtime layer, and optionally emitting WebAssembly. The repo also contains the current CLI/runtime entrypoint, GUI/editor support crates, language examples, and broad per-language regression suites.

## What Is In This Repo

The active workspace currently contains these crates:

```text
crates/
  code_editor/    Code editor and project-facing UI support
  vybe_bytecode/  Bytecode format, VM, and WASM emission support
  vybe_compiler/  Frontends, common AST, lowering, bundling, and tests
  vybe_host/      Host/runtime modules and capabilities
  vybe_widgets/   Shared GUI/widget layer
  vybex/          CLI runner, server mode, and runtime entrypoint
```

The old split parser/compiler crate layout shown in earlier docs is no longer current. Language frontends now live under `crates/vybe_compiler/src/languages/`.

## Supported Frontends

The compiler currently registers these language frontends:

- Visual Basic (`.vb`)
- JavaScript (`.js`)
- Pascal (`.pas`)
- C# (`.cs`)
- Python (`.py`)
- PHP (`.php`)
- Ruby (`.rb`)
- Dart (`.dart`)
- COBOL (`.cbl`)
- Fortran (`.f90`)
- Go (`.go`)

Project loaders currently support:

- `.vybe`
- `.vbproj`
- `.csproj`
- `.pyproj`
- `.ipyproj`

Examples for several languages live under `examples/`, including `vb`, `js`, `python`, `php`, `ruby`, `dart`, `csharp`, `cobol`, `fortran`, `pascal`, and `webroot`.

## Core Capabilities

- Multi-language compilation into a shared bytecode/runtime model
- Shared host/runtime surface across languages
- WASM emission via `--emit-wasm`
- AST and bytecode inspection from the CLI
- Sandboxed execution mode for restricted capabilities
- Portable mode for minimal WASI-style execution
- HTTP directory serving and programmatic server support in `vybex`
- GUI/editor-oriented crates for form and widget workflows

## Build

Build the full workspace:

```bash
cargo build
```

Build optimized binaries:

```bash
cargo build --release
```

Run the CLI directly from the workspace:

```bash
cargo run -p vybex -- path/to/file.js
```

## CLI Usage

The executable name is `vybex`.

```bash
# Compile and run a source file
cargo run -p vybex -- path/to/program.vb
cargo run -p vybex -- path/to/program.js
cargo run -p vybex -- path/to/program.py

# Dump bytecode
cargo run -p vybex -- --dump path/to/program.js

# Dump the prepared common AST
cargo run -p vybex -- --dump-ast path/to/program.php

# Emit WebAssembly
cargo run -p vybex -- --emit-wasm path/to/program.go

# Evaluate source from a string
cargo run -p vybex -- --eval 'print(1 + 2)' --lang python

# Restricted runtime
cargo run -p vybex -- --sandbox path/to/untrusted.py

# Minimal portable runtime
cargo run -p vybex -- --portable path/to/program.js

# Start the directory server
cargo run -p vybex -- --serve --bind 127.0.0.1:8080 examples/webroot
```

Key flags exposed by the current CLI:

- `--dump`
- `--dump-ast`
- `--emit-wasm`
- `--eval CODE --lang NAME`
- `--sandbox`
- `--portable`
- `--trace`
- `--chunk NAME`
- `--serve`
- `--bind ADDR`
- `--no-sandbox`

## Workspace Layout

```text
README.md
Cargo.toml
crates/
examples/
documentation/
data/
tools/
```

- `crates/` holds the compiler, runtime, widgets, editor, and CLI crates.
- `examples/` contains multi-language sample programs and web assets.
- `documentation/` contains architecture notes, plans, and feature docs.
- `data/` contains implementation notes, status docs, and test-planning material.

## Tests

The compiler crate has dedicated per-language regression suites. Common focused commands are:

```bash
# Entire compiler test suite
cargo test -p vybe_compiler

# Single language suites
cargo test -p vybe_compiler --test vb
cargo test -p vybe_compiler --test js
cargo test -p vybe_compiler --test python
cargo test -p vybe_compiler --test php
cargo test -p vybe_compiler --test ruby
cargo test -p vybe_compiler --test dart
cargo test -p vybe_compiler --test csharp
cargo test -p vybe_compiler --test pascal
cargo test -p vybe_compiler --test cobol
cargo test -p vybe_compiler --test fortran
cargo test -p vybe_compiler --test go
```

## Where To Look Next

- `crates/vybex/src/cli.rs` for current CLI behavior and flags
- `crates/vybe_compiler/src/languages/` for language frontends
- `crates/vybe_compiler/tests/` for runnable regression coverage
- `documentation/architecture.md` for higher-level design notes

## License

The workspace package metadata currently declares `GPLv3`.
