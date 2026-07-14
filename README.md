# Vybe

Vybe is a Rust workspace centered on a WebAssembly-oriented VM with an ECMA-shaped runtime layer. It compiles multiple source languages into a shared bytecode model, provides WASI-facing host functionality and proposal-oriented runtime surfaces, and can emit WebAssembly alongside running code inside the native Vybe VM. The repo also contains the current CLI/runtime entrypoint, GUI/editor support crates, language examples, and broad per-language regression suites.

## What Is In This Repo

The active workspace currently contains these key components:

```text
crates/
  code_editor/    Code editor and project-facing UI support
  vybe_bytecode/  Bytecode format, VM, and WASM emission support
  vybe_compiler/  Common AST, lowering, bundling, and tests
  vybe_emitter/   Shared cross-language bytecode helpers
  vybe_host/      Host/runtime modules and capabilities
  vybe_plugin/    Plugin SDK (traits, registry, profiles)
  vybe_widgets/   Shared GUI/widget layer
  vybex/          CLI runner, server mode, and runtime entrypoint
languages/        One crate per language frontend (grammar + walker + emit adapters)
platforms/        One crate per target platform (e.g. dotnet, libc)
```

Language frontends and platforms are now extracted into their own crates under `languages/` and `platforms/`.

## Supported Frontends

The compiler currently registers 15 language frontends:

- C (`.c`, `.h`)
- C# (`.cs`)
- COBOL (`.cob`, `.cbl`, `.cobol`)
- Dart (`.dart`)
- Fortran (`.f90`, `.f95`, `.f03`, `.f08`, `.f18`, `.f`)
- Go (`.go`)
- Java (`.java`)
- JavaScript (`.js`, `.mjs`)
- Lua (`.lua`)
- Pascal / Delphi (`.pas`, `.pp`)
- PHP (`.php`, `.phtml`)
- Python (`.py`, `.pyw`)
- Ruby (`.rb`)
- Visual Basic (`.vb`)
- WebAssembly Text (`.wast`, `.wat`)

Project loaders currently support:

- `.vybe` (Vybe Project)
- `.vbproj` (Visual Basic Project)
- `.csproj` (C# Project)
- `.pyproj` (Python Project)
- `.ipyproj` (IronPython Project)
- `.dpr`, `.lpr` (Delphi / Lazarus Project)

Examples for several languages live under `examples/`.

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

```text
vybex — Universal compiler

Usage: vybex [flags] <file>
       vybex --eval CODE --lang NAME [--virtual-path PATH]
       vybex --serve [--bind BIND] [BIND] [ROOT]

Flags:
  -d, --dump        Disassemble bytecode (no run)
      --dump-ast    Parse and print the prepared common AST
  -w, --emit-wasm   Compile to .wasm binary
      --eval CODE   Compile source from a string
      --lang NAME   Language for --eval (js, php, python, vb, ...)
      --virtual-path PATH  Source path used for relative imports in --eval
  -s, --sandbox     Restricted mode (safe capabilities only)
  -p, --portable    Minimal WASI runtime (no Vybe host)
  -t, --trace       Enable bytecode trace output
      --chunk NAME  Limit --dump/--trace output to a chunk name or index
      --serve       Start HTTP server for a directory (see httpserver.md)
      --bind ADDR   With --serve: bind to ADDR instead of 127.0.0.1:8080
                    BIND defaults to 127.0.0.1:8080, ROOT to current dir
      --no-sandbox  With --serve: keep full host access (default)
  -h, --help        Show this help
```

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

Each language frontend has its own dedicated test suite in its respective crate. Common focused commands are:

```bash
# Single language suites
cargo test -p vybe_language_vb --test vb
cargo test -p vybe_language_js --test js
cargo test -p vybe_language_python --test python
cargo test -p vybe_language_php --test php
cargo test -p vybe_language_ruby --test ruby
cargo test -p vybe_language_dart --test dart
cargo test -p vybe_language_csharp --test csharp
cargo test -p vybe_language_pascal --test pascal
cargo test -p vybe_language_cobol --test cobol
cargo test -p vybe_language_fortran --test fortran
cargo test -p vybe_language_go --test go
```

## Where To Look Next

- `crates/vybex/src/cli.rs` for current CLI behavior and flags
- `crates/vybe_compiler/src/languages/` for language frontends
- `crates/vybe_compiler/tests/` for runnable regression coverage
- `documentation/architecture.md` for higher-level design notes

## License

Vybe is dual-licensed under `GPLv3` and a commercial license.

To be accepted, contributions must be able to ship under either license. By submitting a contribution for inclusion, you are agreeing that the project may distribute that contribution under both licensing tracks.
