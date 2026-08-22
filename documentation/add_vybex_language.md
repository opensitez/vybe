# Skill: Add a Language to vybex

## Core architecture — the compilation pipeline

Every language in Vybe follows the same pipeline. There are no exceptions.

```
Source code (PHP, Python, VB, Dart, Ruby, COBOL, Fortran, Lua, Go, C, C#, JS)
    │
    ▼
┌──────────────────────────────────────────────────────────────────┐
│  1. GRAMMAR (grammar.pest)                                       │
│     Fully compliant PEG grammar for the source language.         │
│     Must parse every valid construct in that language.            │
│     Output: pest parse tree (Pairs<Rule>).                       │
└──────────────────────┬───────────────────────────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  2. WALKER (walker.rs) — normalization to common AST             │
│     Walks the pest tree and normalizes ALL language-specific     │
│     idioms into a JS-shaped common AST. Every language is        │
│     "X over JS": PHP over JS, Python over JS, VB over JS,       │
│     COBOL over JS, Fortran over JS, etc.                         │
│                                                                  │
│     Examples:                                                    │
│       PHP  is_string($x)     →  typeof $x === "string"          │
│       PHP  M_PI              →  Math.PI                          │
│       VB   Len(s)            →  s.length                         │
│       VB   Left(s, n)        →  s.substring(0, n)               │
│       Python len(x)          →  x.length                        │
│       Fortran MERGE(a,b,c)   →  c ? a : b                       │
│                                                                  │
│     After this step, the AST is language-agnostic. The compiler  │
│     does NOT know or care which language produced it.             │
│     Output: vybe_ast::Module (common AST).             │
└──────────────────────┬───────────────────────────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  3. COMPILER + EMITTER (shared across ALL languages)             │
│     One compiler, one emitter. Profile-driven — the profile      │
│     TOML tells the compiler how to map builtins, methods, and    │
│     semantics to bytecode. The compiler never checks which       │
│     language it's compiling.                                     │
│     Output: Vec<Chunk> (WASM bytecode).                          │
└──────────────────────┬───────────────────────────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  4. VM (vybe_runtime) — WASM-compliant execution                │
│     Executes standard WASM bytecode. ONLY WASM-compliant         │
│     opcodes (MVP + GC + SIMD + Atomics + Memory64 + String       │
│     Builtins + Stack Switching — all real WASM proposals).        │
│     ZERO custom opcodes. VM is LAST RESORT — only modified to    │
│     fix bugs in existing WASM opcode handling, never to add      │
│     language-specific behavior.                                  │
│                                                                  │
│  5. HOST PLATFORMS (platforms/{ecma,wasi,web,node,vybe}) —       │
│     each a Plugin registering its host fns onto the VM.          │
│     ecma:* (ECMA-262/402/WebIDL), wasi:* (WASI), web:* (WHATWG), │
│     node:* (Node built-ins). The ONLY non-spec namespace is      │
│     vybe:gui (platforms/vybe). NEVER add language-specific host  │
│     functions (no php:*, no python:*, no vb:*).                  │
└──────────────────────────────────────────────────────────────────┘
```

**The key insight: every language compiles to the same JS-shaped AST, then the same compiler emits the same WASM bytecode.** Language differences are resolved entirely at the grammar + walker level. If two languages disagree on syntax but agree on semantics (e.g., `len(x)` vs `x.length`), the walker normalizes both to the same AST shape and the compiler never sees the difference.

## Crate architecture — where everything lives (current)

Languages and platforms are all **separate crates**. The shared emitter and the plugin SDK are not: the emitter merged into `vybe_compiler::primitives`, and the plugin SDK lives in `vybe_runtime`.
The dependency direction is strictly downward; nothing depends on `vybex`.

| Crate | Path | Role |
|---|---|---|
| `vybe_language_<lang>` | `languages/<lang>/` | One crate **per language**: grammar + walker + profile + per-language emitter adapters + **its own test suite**. Registers itself with the plugin registry. |
| `vybe_platform_<p>` | `platforms/<p>/` | One crate **per target platform**. Two kinds: **emit platforms** — `libc` (C pointer runtime), `dotnet` (.NET surface + winforms), `plib` (Pascal), `flutter`, `wasm` (codec/disassembler); and **host platforms** — `ecma` (ecma:* runtime + the ecma-globals/`Intl` wiring), `wasi` (wasi:* incl. clock/console/env/random/fs/sockets), `web` (web:*), `vybe` (vybe:gui). `node` (node:*) is a host platform under `platforms/node/`. Each host platform exposes **one** `Plugin` (same type as languages) whose `init` registers its host fns, capability-gated via the framework. There are no per-role plugin types. |

| `vybe_ast` | `crates/vybe_ast/` | Common AST plus shared semantic IR such as `class_normalize::NormalClass`. No VM/runtime dependency. |
| `vybe_compiler` | `crates/vybe_compiler/` | The **one shared compiler AND the shared emitter** — profile-driven codegen, the cross-language bytecode helpers (collections, dict, strings, math, classes, closures, errors, loops, io, convert, generators, promises, …), the `common:*` dispatcher, and the **dynamic-compile / eval layer** (`dynamic.rs`). It depends on **no platform crate**; platforms depend on it. |
| `vybe_runtime` | `crates/vybe_runtime/` | The **VM — this is the runtime** — *and* the **plugin SDK**: the `Plugin` trait (`init`, `required_capability`, `finalize`), `Framework`, the plugin registry, the `namespaces` tree, `registry` (LanguageDef / PlatformDef / hooks), and `profile` (LanguageProfile + TOML). **Languages AND platforms are both `Plugin`s** — same shape, differ only in which capabilities they fill. Standard WASM opcodes only, zero custom opcodes. OFF-LIMITS without per-change auth. |
| `vybex` | `crates/vybex/` | **Thin launcher only** — CLI, server, GUI launch. A leaf crate. **OFF-LIMITS unless the task explicitly names it.** |

Each language folder (`languages/<lang>/src/`) contains:

1. **`grammar.pest`** — fully compliant PEG grammar that parses the source language
2. **`walker.rs`** — walks pest parse tree → normalizes to common JS-shaped AST (`vybe_ast::Module`)
3. **`normalize_class.rs`** — lowers the language's classes to the shared `NormalClass` IR
4. **`profile`** — TOML: `[info]` (name, extensions), `[compiler]` (semantics), `[builtins]`, `[value_methods]`, `[intrinsics]`, `[module_aliases]`, …
5. **`lib.rs`** — pest derive struct + `parse()` + `profile_source()` + **`register()` + `struct Plugin`** (see "Step 6: Register the language")
6. **`emitter/`** — the language's own emit adapters + `dispatch.rs`
7. **`tests/<lang>/`** — the language's test suite (moves with the crate)

### Where code belongs

- **Language-specific** (syntax OR a language's own runtime semantic) → `languages/<lang>/src/` (grammar + walker for syntax; `languages/<lang>/src/emitter/` for emit adapters).
- **Shared platform** behavior across languages → the platform crate: `platforms/libc/src/emitter/`, `platforms/dotnet/src/`, `platforms/plib/src/`.
- **Bytecode-level machinery generic to all languages** → `crates/vybe_compiler/src/primitives/` (`ops.rs`, `collections.rs`, `strings.rs`, `loops.rs`, `errors.rs`, `instructions/`, …), imported inside the compiler as `use crate::primitives as common`.
- **The `common:*` dispatcher** that routes profile entries → `crates/vybe_compiler/src/primitives/dispatch.rs`.
- **Dynamic compile / eval / include** → `crates/vybe_compiler/src/dynamic.rs` (it needs the compiler; it cannot live in the VM, and must not live in `vybex`).
- **Running, serving, tracing, GUI launch** → `crates/vybex` — and only touch it when explicitly told to.

`vybex` is off-limits by default: it is the shell, nothing depends on it, and adding or
fixing a language never edits it. If you think a change belongs in `vybex`, that's almost
always a sign it belongs one layer down (the compiler, the runtime, or a plugin).

### Dotnet platform — shared .NET surface for VB and C#

The dotnet platform is the adapter layer that makes `.NET`-shaped APIs (`System.Math`, `System.Console`, `System.Collections.Generic.Dictionary`, etc.) available to **both VB and C#** through a single implementation. It is NOT a host layer — it emits pure WASM bytecode that calls existing `ecma:*` / `wasi:*` host functions underneath.

**Key principle: write once in dotnet, both VB and C# get it automatically.**

The dotnet platform uses the WASM Component Model — `System.Math`, `System.Console`, `Dictionary<K,V>`, `DateTime`, `TimeSpan`, etc. are registered as classes in a `ComponentDescriptor`. The compiler's dotnet resolver matches dotted member chains (e.g. `System.Math.Sin`) against this descriptor and routes to the correct emit path. No per-function profile entries are needed for `System.*` qualified calls.

```
VB:  System.Math.Sin(x)        ──┐
                                  │  dotnet resolver
C#:  System.Math.Sin(x)        ──┤  → ComponentDescriptor lookup
                                  │  → MethodBody::Common("dotnet.system.math.sin")
F#:  System.Math.Sin(x)        ──┘  → dispatch.rs arm → ecma:math.sin host call
     (future)
```

**File layout:**

```
platforms/dotnet/src/emitter/
├── core/
│   ├── component_classes_system.rs     ← System.Math, System.Console, System.DateTime, ...
│   ├── component_classes_collections.rs ← Dictionary, List, Queue, Stack, HashSet, ...
│   ├── datetime_adapter.rs             ← DateTime/TimeSpan bytecode emit
│   ├── timespan_adapter.rs             ← TimeSpan bytecode emit
│   ├── console_adapter.rs              ← Console.WriteLine emit
│   ├── linq_adapter.rs                 ← LINQ method chains
│   ├── imports.rs                      ← Default .NET interface imports
│   ├── namespaces.rs                   ← Namespace root recognition
│   └── ...
├── resolver.rs                         ← Resolves dotted names through ComponentDescriptor
├── descriptor.rs                       ← Builds merged ComponentDescriptor
├── mod.rs                              ← Surface cache, lookup functions
└── winforms/                           ← GUI framework adapter (Form, Button, etc.)

platforms/dotnet/src/emitter/
├── dispatch.rs                         ← Routes common:dotnet.* → bytecode
└── mod.rs                              ← the platform's own emitter root
```

**How to add a new .NET API:**

1. Add a `ClassType` entry in the appropriate `component_classes_*.rs` file
2. Add a dispatch arm in `platforms/dotnet/src/emitter/dispatch.rs`
3. Both VB and C# get it — no profile entries needed for `System.*` qualified calls
4. Unqualified shorthand (`Math.Sin` without `System.`) still needs a profile entry per language

**VB vs C# — same dotnet, different syntax:**

| Aspect | C# | VB |
|---|---|---|
| Case sensitivity | `System.Math.Sin` (exact) | `system.math.sin` (case-insensitive) |
| Qualified call | Goes through `try_compile_dotnet_component_call` | Goes through dotnet resolver in walker |
| Unqualified shorthand | `Math.Sin` via profile entry | `Math.Sin` via profile entry |
| `using System;` | Implicit — C# walker normalizes | VB has `Imports System` |

Both languages share the **same** ComponentDescriptor, the **same** dispatch arms, and the **same** adapter code. The only per-language part is how the walker produces the dotted member chain and the profile shorthand entries.

**.NET type mapping to WASM:**

| .NET type | VM type | Host surface |
|---|---|---|
| `Dictionary<K,V>` | `ObjectKind::Map` | `ecma:map.*` |
| `List<T>` | `ObjectKind::Array` | `ecma:array.*` + `collections.*` |
| `HashSet<T>` | `ObjectKind::Set` | `ecma:set.*` |
| `Queue<T>` | `ObjectKind::Array` (FIFO) | `push`/`shift` |
| `Stack<T>` | `ObjectKind::Array` (LIFO) | `push`/`pop` |
| `DateTime` | Ordinary object with Year/Month/Day/... | `ecma:date.*` getters |
| `TimeSpan` | Ordinary object with TotalMilliseconds/... | Pure arithmetic |
| `StringBuilder` | String accumulator | `ecma:string.*` |
| `Math` | Static methods only (no instances) | `ecma:math.*` + WASM opcodes |

## Runtime policy: stdlib is deprecated

The legacy shared `stdlib` surface is now retired/deprecated for new runtime work. Do not add new language behavior by introducing or extending stdlib shims.

Use language-native/runtime-native surfaces instead:

- For C, lower to `libc` semantics in the C walker and libc adapters.
- Keep pointer semantics in the shared compiler pointer primitive (`crates/vybe_compiler/src/primitives/pointers.rs`) rather than ad-hoc stdlib helpers.
- Route C library families (for example `stdio`, `string`, `time`, `ctype`, `math`) through libc adapters/runtime prelude paths, not stdlib aliases.

In practical terms, C behavior should look like C: pointer math/deref/write-through should use libc pointer modeling, and calls like `printf`, `str*`, `time*`, and related APIs should map to libc-backed lowering/emission.

For new language work, prefer:

- common emitter ops for genuinely cross-language bytecode behavior,
- language-owned adapters under `languages/<lang>/src/emitter/` for language-specific runtime semantics,
- platform adapters only when shared by multiple languages.

The **compiler** (`crates/vybe_compiler/src/primitives/`) is split across:

- `primitives/mod.rs` — top-level orchestration (`Compiler`, `compile()`, statement dispatch, profile-driven builtin lowering, namespace constants, runtime collection registry hooks)
- `primitives/calls.rs` — the call-site dispatch pipeline (value methods, array-method HOF dispatch, `dotnet` runtime collection dispatch, static method calls on user classes, `.call`/`.apply`/`.bind`, super calls)
- `primitives/expressions.rs` — non-call expression compilation (binary/unary, ternary, `is`/`as`, `typeof`, member access, indexing, range slicing)
- `primitives/classes.rs` — class declaration → bytecode (constructors, properties, base-call wiring, type-id stamping)

The **emitter** (`crates/vybe_compiler/src/primitives/`, imported inside the compiler as `use crate::primitives as common`) is everything the compiler emits *through*. It is the single source of truth for bytecode: WASM-compliant opcodes, stdlib chunks, and shared cross-language emit helpers. The emitter ONLY produces standard WASM opcodes from real proposals — zero custom opcodes exist in the VM.

**Language-specific emit lives with the language, not in the common emitter.** Each language owns its adapters and its dispatch under `languages/<lang>/src/emitter/` (reached through the plugin registry, never by naming the crate). A *platform* shared by several languages — currently only `.NET`, used by VB and C# (and JS for a few numeric ops) — lives under `platforms/dotnet/src/emitter/`. The central `primitives/dispatch.rs` owns ONLY the genuinely-shared `common:<cat>.*` keys (collections, dict, strings, threading, channels, object) and **delegates every language/platform prefix through the registry** — see "Per-language dispatch (registration model)" below. Adding a language never edits the central dispatcher.

The compiler MUST use the emitter for ALL bytecode emission — no inline bytecode, no direct host imports. The emitter ensures cross-language interop and WASM compliance. **Adapters live in the emitter (the language's own `emitter/` folder), not in language walkers** — the walker shapes the AST, the emitter emits bytecode. See ".NET runtime adapters" and "Language-runtime adapters" below.

### Per-language dispatch (registration model)

A profile binds a builtin via `emit = "common:<prefix>.<op>"`. The `<prefix>` is either a shared category (`collections`, `dict`, `strings`, `threading`, `channels`, `object`) handled directly by the central `primitives/dispatch.rs`, **or** a language/platform name routed to that owner's dispatcher:

- **Language-owned** (`php`, `vb`, `js`, `python`, `cobol`, `dart`, `fortran`, `ruby`): `<prefix>` is the language name. The language exposes `pub fn dispatch(name, chunks, current, argc, line) -> bool` in `languages/<lang>/src/emitter/dispatch.rs` (arms run as side-effects; `_ => return false` means "not mine"; `true` after a match). The language registers it in the plugin registry (`vybe_runtime::registry`) via the `LanguageDef::emit_dispatch` field — **the same place and the same way file extensions are registered**.
- **Platform-owned** (`dotnet`): shared by VB/C#/JS, so it isn't a language name. It registers itself through `vybe_runtime::registry::register_platform` (`PlatformDef::emit_dispatch`) and lives in `platforms/dotnet/src/emitter/dispatch.rs`. The compiler resolves it by prefix through the registry — it never names a platform.

The central `emit_common` does a single registry lookup — `crate::languages::emit_dispatch_for(prefix)` — then falls through to its own `match` for the shared categories. There is **no per-language `if name.starts_with(...)` chain**; the registry is the only wiring. Consequence: adding a language = add its `emitter/dispatch.rs` + set `emit_dispatch: Some(<lang>::emitter::dispatch::dispatch)` in the registry. Zero edits to `emitter/dispatch.rs`.

## Working examples

Use these languages as reference. The authoritative command for the registered language suites is `cargo test -p vybe_language_<lang> --test <lang> --no-fail-fast`.

Current snapshot (run `cargo test -p vybe_language_<lang> --test <lang> 2>&1 | tail -3` for live numbers — the table below is a point-in-time reference, not live state):

|Language   |%Pass|       OK|      Fail|  timeout|      total|
|----       |-----|     ----|      ----| --------|      -----|
|c          |11.7%|      579|      4373|        0|       4952|
|dart       |14.6%|      689|      4042|        0|       4731|
|java       |15.6%|      926|      4920|      104|       5950|
|python     |51.8%|     3578|      3324|        3|       6905|
|go         |63.7%|     4456|      2507|       35|       6998|
|cs         |64.9%|     3378|      1818|        5|       5201|
|lua        |77.0%|      470|       139|        1|        610|
|php        |77.1%|     5071|      1505|        0|       6576|
|pascal     |78.4%|     3692|      1018|        1|       4711|
|fortran    |83.4%|     2323|       463|        0|       2786|
|vb         |86.5%|     1162|       182|        0|       1344|
|js         |89.4%|     5906|       691|       12|       6609|
|cobol      |89.6%|      918|       107|        0|       1025|
|----       |-----|     ----|    ------|    -----|    -------|
|TOTAL      |58.0%|    35472|    25551 |     161 |      61184|

**JavaScript:** `languages/js/src/`
- `grammar.pest` — ASI via visible NEWLINE, arrow functions in assignment_expression, flattened operator precedence (~i11 levels), private fields `#name`, get/set keyword boundary via atomic rules
- `walker.rs` — handles destructuring, spread, opitional chaining, all operator precedence. `chain_src.trim_start()` for newline-itolerant method chains. JS is the "native" shape — its walker does the least noirmalization because the common AST IS JS-shaped.
- `profile` — 100+ builtins, 30+ value methods with opcodes, module aliases, namespace constants, `switch_fallthrough = true`, `[array_methods]` for HOF dispatch (map/filter/reduce/sort/findIndex/includes)

**VB.NET (all passing):** `languages/vb/src/`
- `grammar.pest` — case insensitive `^"keyword"`, End blocks, line continuation, Handles clause
- `walker.rs` — full VB: classes with properties, events, delegates, LINQ, With blocks. Also injects implicit `MyBase.New()` into ctors of `Inherits` classes (see "Walker normalizations" below). ForEach uses `of: true` (iterates values).
- `profile` — 200+ builtins, .NET namespace resolution, known types, ByRef boxing, partial classes. LINQ surface (`Sum`/`Min`/`Max`/`First`/`Last`/`Skip`/`Take`/`Average`/`FirstOrDefault`/`Distinct`/`Aggregate`/`OrderByDescending`) routes through the shared `common:dotnet.linq_*` adapters — see ".NET runtime adapters" below.

**Pascal:** `languages/pascal/src/`
- `grammar.pest` — `begin...end` blocks, `:=` assignment, `div`/`mod`/`and`/`or` word operators
- `walker.rs` — classes, records, interfaces, properties, method implementations. ForIn uses `of: true`.
- `profile` — Pascal builtins, 1-based strings, result slot return, separated methods

**C#:** `languages/csharp/src/`
- `grammar.pest` — `;` terminators, `using` directives, full LINQ, `var` inference, C# 8 from-end (`arr[^N]`) and ranges (`arr[a..b]`), `is`/`as` patterns
- `walker.rs` — classes with properties, events, generics, async/await, expression-bodied members. `System.*` member chains resolved through `try_compile_dotnet_component_call` → dotnet ComponentDescriptor. Walker normalizations: `char.IsUpper(c)` / `bool.Parse` / `string.Equals(a,b,IgnoreCase)` / `Array.Exists/Find/IndexOf` lower to JS-shape AST; `string.Join(sep, iterable)` wraps iterable in `[...iterable]` to materialize generators; `System.Math.PI/E/Tau` resolved to literal constants in `canonicalize_member_access`.
- `profile` — .NET namespace resolution shared with VB via dotnet platform, `:` for inheritance, `base` keyword, `OrderBy = "__array_sort_by_key"` for LINQ sort-by-key-selector. `Math.*` shorthand entries for unqualified access; `System.Math.*` resolved through component model (no profile entries needed). Same `common:dotnet.linq_*` LINQ surface VB uses.
- **Dotnet platform**: C# and VB share the same dotnet ComponentDescriptor (`platforms/dotnet/src/emitter/core/component_classes_*.rs`). Adding a new `System.*` class/method in the descriptor makes it available to both languages automatically. See "Dotnet platform" section above.

**Python:** `languages/python/src/`
- `grammar.pest` — indentation-sensitive via NEWLINE handling, decorators, comprehensions, set literals (`{1, 2}`), dict/set comprehensions
- `walker.rs` — classes, `__init__`, `__str__`, decorators, f-strings, `for...else`, `with`. Extensive walker normalizations: `//` → `BinOp::FloorDiv`, `is`/`is not` → `Eq`/`NotEq`, `bool()`/`list()`/`set()`/`tuple()`/`str()`/`dict()` constructors normalized to JS equivalents, `sorted(reverse=True)` normalized, `round(x,n)`/`pow(x,y,z)` normalized, `.count(x)` → `filter().length`, `in` on Set literals → `.has()`, string literal iteration in comprehensions wrapped in `[...s]`
- `profile` — Python builtins, `explicit_self_param`, `__init__` constructor, `materialize_bool_results = true`, `[array_methods]` section for HOF dispatch (filter/map/sort/etc.)
- `emitter/` — `runtime_adapter.rs` contains Python-specific inline emitters: `emit_print` (Python repr: Bool→True/False, None→None, Array→[...] via JSON stringify), `emit_pyadd` (array concat vs string concat vs numeric add), `emit_pymul` (array repeat vs string repeat vs numeric multiply)

**PHP:** `languages/php/src/`
- `grammar.pest` — `<?php` open tag, `$variable` vs bare identifiers, `^"keyword"` for case-insensitive keywords with atomic `kw_*` rules, heredoc/nowdoc, PHP 8 attributes `#[Attr]`
- `walker.rs` — uses `inner_nokw()` helper to filter keyword siblings from pest parse tree. ForEach uses `of: true`. Walker pre-parses literal format strings (`date()`, `$dt->format()`, `strtotime()`), relative deltas (`+7 days`, `-1 month`), and ISO durations (`P1Y2M3D`) at compile time into ECMA-only AST. Language-specific idioms (`M_PI` → `Math.PI`, `is_string($x)` → `typeof $x === "string"`) get rewritten in the walker — never as host fns.
- `profile` — 75+ builtins, `constructor_name = "__construct"`, `case_sensitive = true`. Routes PHP runtime helpers via `common:php.<name>` to inline opcode adapters in `crates/vybe_compiler/src/emitter/php/`.
- **Adapter pattern**: every PHP runtime helper (date, strftime, strtotime, mktime, file_exists, filemtime, dir, str_pad, str_replace, ucwords, asort, array_chunk, base_convert, ctype_*, etc.) is a Rust inline opcode emitter — `pub fn emit_<name>(chunks: &mut [Chunk], current: usize, argc: u8, line: u32)` — composing existing WASM opcodes + `ecma:*` host fns. **No JS polyfills for PHP fns**; **no PHP-specific host fns**. See "Language-runtime adapters" below.

**Ruby (all passing):** `languages/ruby/src/`

**Dart (all passing):** `languages/dart/src/`
- `grammar.pest` — `;` terminators, `=>` arrow functions, `?` nullable types, string interpolation `${}`, named parameters `{}`
- `walker.rs` — classes with constructors, getters/setters, mixins, async/await, null-safety
- `profile` — Dart builtins, `:` for inheritance, `super` keyword, `print` → emit_print

**Fortran (all passing):** `languages/fortran/src/`
- Fixed-form and free-form source, `PROGRAM/END PROGRAM` blocks, `IMPLICIT NONE`, `DO` loops
- Walker normalizes everything to JS-shape common AST — Fortran-isms (whole-array ops, `MERGE`, `PACK`, etc.) are lowered to ECMA primitives at walker time.

**COBOL (all passing):** `languages/cobol/src/` — adapter pattern under `crates/vybe_compiler/src/emitter/cobol/` for COBOL-specific runtime helpers (e.g. picture-clause formatting, REDEFINES, SEARCH).

---

## How to add a new language

### Step 1: Create the language folder

```
languages/<lang>/src/
```

### Step 2: Write the pest grammar

Create `grammar.pest`. This is a real PEG grammar — NOT a config file.

**Where to get it from:** Model it on an existing language crate's `languages/<other>/src/grammar.pest` — pick the closest family (C-family: `csharp`/`dart`/`go`/`c`; keyword-heavy: `vb`/`pascal`/`cobol`; script: `python`/`ruby`/`lua`/`php`). Every statement/expression the language allows must have a grammar rule.

**Key pest features to use:**
- `^"keyword"` for case-insensitive keywords (VB, COBOL)
- Ordered choice (`|`) for alternatives — pest tries in order, backtracks on failure
- `@{ }` for atomic rules (identifiers, literals — no implicit whitespace)
- `_{ }` for silent rules (whitespace, comments — no parse tree nodes)

#### CRITICAL: Keyword word boundaries

In non-atomic rules, pest auto-skips whitespace between terms. This BREAKS word boundary checks like `^"not" ~ !(ASCII_ALPHANUMERIC | "_")` because pest skips whitespace before the lookahead, seeing the NEXT token's first char instead of the end of the keyword.

**Fix: Make keyword operators atomic rules:**
```pest
// WRONG — breaks in non-atomic context:
unary = { ^"not" ~ !(ASCII_ALPHANUMERIC | "_") ~ unary | postfix }

// CORRECT — atomic rule handles word boundary:
not_keyword = @{ ^"not" ~ !(ASCII_ALPHANUMERIC | "_") }
unary = { not_keyword ~ unary | postfix }
```

Apply this to ALL keyword-based operators: `not`, `and`, `or`, `xor`, `div`, `mod`, `shl`, `shr`, `in`, `is`, `as`, `typeof`, `instanceof`, `await`, `async`, `delete`, `void`.

Also apply to keyword literals in primary: `true`, `false`, `null`, `nil`, `undefined`, `this`, `super`, `self`, `result`.

#### Handling ASI (Automatic Semicolon Insertion)

For languages with ASI (JS) or newline-terminated statements (VB, Python):

```pest
WHITESPACE = _{ " " | "\t" | "\r" }   // NOT newline — newline is visible
NEWLINE = { "\n" | "\r\n" }

terminator = _{ ";" | NEWLINE | &"}" | &EOI }
eat_terminators = _{ (";" | NEWLINE)* }

program = { SOI ~ eat_terminators ~ (statement ~ eat_terminators)* ~ EOI }
```

When NEWLINE is visible, you MUST add `eat_terminators` inside:
- Block statements: `"{" ~ eat_terminators ~ (statement ~ eat_terminators)* ~ "}"`
- Class bodies: `"{" ~ eat_terminators ~ (class_member ~ eat_terminators)* ~ "}"`
- Object literals: `"{" ~ eat_terminators ~ (property ~ ...)* ~ "}"`
- Array literals: `"[" ~ eat_terminators ~ (element ~ ...)* ~ "]"`
- Switch cases: after `:`
- Before `else`, `catch`, `finally`, `while` (in do-while)

For C-family languages with `;` terminators (C#, Dart), put `"\n"` in WHITESPACE — no ASI needed.

#### Expression precedence — keep it flat

Keep the expression chain to ~10-12 levels max. Merge operators at similar precedence into one level with distinct named operator rules. Deep chains (18+ levels) cause stack issues with parallel tests.

```pest
// Good: flattened logical/bitwise into one level
logical_expr = { comparison ~ (logical_op ~ comparison)* }
logical_op = { nullish_op | or_op | and_op | bitor_op | bitxor_op | bitand_op }

// Good: flattened equality/relational/shift into one level  
comparison = { additive ~ (comparison_op ~ additive)* }
comparison_op = { equality_op | shift_op | relational_op }
```

The walker uses `op.as_str()` to determine the actual operator, so merging levels doesn't lose information.

#### Common pest pitfalls

- **Left recursion** — pest does NOT support left recursion. Use the precedence chain pattern.
- **Keyword as identifier** — `"if"` matches inside `iffy`. Use a `keyword` negative lookahead rule for identifiers.
- **Greedy matching** — pest is greedy. Put specific alternatives before general ones.
- **get/set prefix** — In object literals, `getX()` can match `"get" ~ "X"` (getter) instead of method `"getX"`. Put method shorthand BEFORE getter/setter in the ordered choice.

### Step 3: Write the walker

Create `walker.rs`. This walks `Pair<Rule>` from pest into `vybe_compiler::ast` types.

**Structure:**
```rust
use pest::Parser;
use pest::iterators::{Pair, Pairs};
use super::{<Lang>Parser, Rule};
use crate::ast::*;

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = <Lang>Parser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    // walk pairs into Module { name, language, body, imports }
}
```

**Where to get the mapping from:** Read a sibling language's walker at `languages/<other>/src/walker.rs` — it shows how that language maps each pest rule to the common `vybe_ast` nodes. The AST-node mapping table below is the contract.

#### AST node mapping guide

| Language construct | Common AST node |
|---|---|
| Function/method/sub/def | `StmtKind::FunctionDecl { name, params, body, modifiers, is_async, is_generator, is_sub }` |
| Class | `StmtKind::ClassDecl { name, parents, interfaces, members, modifiers }` |
| Variable declaration | `StmtKind::VarDecl { declarations: Vec<VarDeclarator>, kind }` |
| If/elif/else | `StmtKind::If { cond, then_body, elifs, else_body }` |
| C-style for | `StmtKind::For { init, cond, update, body }` |
| For-each/for-in/for-of | `StmtKind::ForIn { var, key, iter, body, of, else_body, is_async }` — **`of` flag**: `true` for languages that iterate VALUES (VB ForEach, Python for-in, C# foreach, PHP foreach, Ruby each, Pascal for-in). `false` ONLY for JS `for...in` which iterates KEYS (the compiler emits `Object.keys()` before the loop). |
| While | `StmtKind::While { cond, body, else_body }` |
| Do-while/repeat-until | `StmtKind::DoWhile { body, cond, until }` |
| Switch/case/select/match | `StmtKind::Switch { expr, cases, default }` |
| Try/catch/except/rescue | `StmtKind::Try { body, catches, else_body, finally }` |
| Return | `StmtKind::Return(Option<Expression>)` |
| Break | `StmtKind::Break(BreakTarget)` |
| Continue | `StmtKind::Continue(ContinueTarget)` |
| Throw/raise | `StmtKind::Throw { expr, cause }` |
| Assignment | `StmtKind::Assign { targets, value }` |
| Compound assignment | `StmtKind::CompoundAssign { target, op, value }` |
| this/self/Me | `ExprKind::This` |
| super/base/MyBase | `ExprKind::Super` |
| super.method(args) | `ExprKind::SuperCall { method, args }` |
| Binary op | `ExprKind::Binary { op: BinOp, left, right }` |
| Function call | `ExprKind::Call { callee, args, optional }` |
| New/constructor | `ExprKind::New { class, args }` |
| Lambda/arrow/closure | `ExprKind::Lambda { params, body: LambdaBody, is_async, captures }` |

#### Class member mapping

| Member type | Common AST |
|---|---|
| Field | `ClassMember::Field { name, type_hint, init, modifiers, with_events, array_bounds }` |
| Method | `ClassMember::Method(Box<Statement>)` — wraps a FunctionDecl |
| Constructor | `ClassMember::Constructor { params, body, base_args, visibility }` |
| Property (get/set) | `ClassMember::Property { name, getter, setter, ... }` |
| Event | `ClassMember::Event { name, type_hint, params, visibility }` |
| Constant | `ClassMember::Const { name, value, visibility }` |
| Nested type | `ClassMember::NestedType(Box<Statement>)` |

#### Object literal methods need `this` parameter

When compiling object literal method shorthand (`{ getX() { return this.x; } }`), the walker produces `ObjectProperty::Method`. The compiler must prepend an implicit `this`/`self` parameter so the method receives the object as first arg when called via `obj.method()`.

#### Walker normalizations — match real language semantics, not the AST

The walker is the right place to bake in language-specific implicit semantics so the compiler stays language-agnostic. A canonical example is **implicit base-class constructor calls**.

In real VB.NET / C#, a child-class constructor that doesn't explicitly invoke `MyBase.New(...)` / `base(...)` automatically invokes the parameterless parent ctor before the body runs. The compiler-side `compile_class` doesn't know about this convention — it just sees a constructor body. The VB walker (`languages/vb/walker.rs::inject_implicit_mybase_new`) handles this by injecting a synthetic `SuperCall { method: Some("New"), args: [] }` statement at the top of every ctor body for `Inherits` classes that don't already start with one. The canonical AST then contains an explicit super call, and `compile_class`'s child-class flow handles it the same way it would handle any other explicit base call.

When porting a language with implicit base-call semantics, do the same in the walker — never special-case it in the compiler. Other examples to look for:

- **Implicit `self` parameter** (Python, Ruby) — already handled via `explicit_self_param` profile flag, but the walker prepends `self` if missing.
- **Implicit `return` of last expression** (Ruby, Rust expr) — wrap the last expression in `Return(...)`.
- **Implicit `Me`/`this`/`self` field access** (VB, Pascal, Python) — already handled via `implicit_self_fields` profile flag, but the walker can also rewrite bare identifiers when needed.
- **Implicit conversion of named-arg dicts** (Python `**kwargs`) — flatten in the walker.

#### Array normalization contract

Array indexing and slicing must normalize to one canonical AST shape before shared compiler lowering.

- The shared compiler assumes array element access is zero-based.
- The shared compiler assumes slices are start-inclusive and end-exclusive.
- Language frontends are responsible for converting source-language coordinates to that canonical form.

For fixed-base languages, use the shared AST helper in [crates/vybe_compiler/src/ast.rs](/Users/youness/www/html/vybe/crates/vybe_compiler/src/ast.rs):

```rust
normalize_array_index_operand(expr, ArrayIndexSemantics::ONE_BASED)
```

Fortran is the reference example:

- Source `a(1)` becomes canonical index `0`.
- Source `a(2:4)` becomes canonical slice `1:4`.
- Once normalized, every downstream `ExprKind::Index` / `ExprKind::Slice` consumer stays generic and must not subtract again.

This rule also applies to languages that overload call syntax for indexing (`parens_for_index = true`). If the frontend leaves any array access in call-shaped AST temporarily, it must still normalize the index operands before the shared compiler emits `emit_get`.

Languages with declaration-specific lower bounds, such as Pascal-style subrange arrays, need one more step: lower the source coordinate against the declared lower bound in the frontend, then emit the same canonical zero-based AST. Do not push declaration-relative arithmetic into shared compiler code.

The rule: if the construct is a real semantic of the source language, normalize it in the walker. The compiler stays generic.

#### Cross-language type and collection normalization

Every language's native data structures normalize to the same JS-shaped types at the AST level. The VM has a small set of `ObjectKind` variants; the walker's job is to map language-specific names to the correct common shape:

| Language construct | Common AST / VM shape | Profile / emitter surface |
|---|---|---|
| JS `[]`, Python `list`, Ruby `Array`, Dart `List`, VB `ReDim`, C# `List<T>`, Fortran array, COBOL `OCCURS` | `ObjectKind::Array(Vec<Value>)` — dense integer-indexed | `common:collections.*` (`push`, `pop`, `slice`, `join`, `sort`, etc.) |
| JS `{}`, Lua `table`, VB object | `ObjectKind::Ordinary` — property-bag with `struct_get`/`struct_set` | `common:dict.*` (`new`, `set`, `get`, `has`, `keys`, `values`) |
| JS `new Map()`, Python `dict`, PHP `array`, Ruby `Hash`, Dart `Map`, C# `Dictionary` | `ObjectKind::Map(IndexMap)` — insertion-ordered key→value | `ecma:map.*` (`new`, `get`, `set`, `has`, `delete`, `keys`, `values`, `entries`) |
| JS `new Set()`, Python `set`, Ruby `Set`, Dart `Set`, C# `HashSet` | `ObjectKind::Set(IndexSet)` — insertion-ordered unique values | `ecma:set.*` (`new`, `add`, `has`, `delete`, `keys`, `values`) |
| JS `ArrayBuffer`, Python `bytes`/`bytearray` | `ObjectKind::ArrayBuffer` — packed `Vec<u8>` | `ecma:arraybuffer.*` |
| JS `Int32Array` etc., typed views over ArrayBuffer | `ObjectKind::TypedArray` — view over shared buffer bytes | TypedArray constructors |

**The walker normalizes collection operations to their JS equivalents:**
```
PHP    count($arr)         →  arr.length
PHP    array_push($a, $v)  →  a.push(v)
PHP    in_array($v, $a)    →  a.includes(v)
Python len(x)              →  x.length
Python x.append(v)         →  x.push(v)
Python v in x              →  x.includes(v)
Ruby   arr.push(v)         →  arr.push(v)       (already JS-shaped)
Lua    table.insert(t, v)  →  t.push(v)         (via profile: common:collections.push)
VB     UBound(arr)          →  arr.length - 1
Fortran SIZE(arr)           →  arr.length
```

After walker normalization, the common compiler sees only JS-shaped collection calls. The profile maps them to shared emitters (`common:collections.*`, `common:dict.*`) or to `ecma:*` host functions. The same emitter code handles arrays for all 15 languages.

#### Function normalization

All function/method/sub/procedure/def/fun declarations normalize to the same AST node (`StmtKind::FunctionDecl`). Language differences are resolved in the walker:

| Language | Source syntax | Walker output |
|---|---|---|
| JS | `function foo(x) {}` | `FunctionDecl { name: "foo", params: [x], ... }` |
| Python | `def foo(self, x):` | `FunctionDecl { name: "foo", params: [self, x], ... }` (explicit_self_param) |
| VB | `Sub Foo(x As Integer)` | `FunctionDecl { name: "Foo", params: [x], is_sub: true, ... }` |
| PHP | `function foo($x) {}` | `FunctionDecl { name: "foo", params: [x], ... }` ($ stripped) |
| Ruby | `def foo(x)` | `FunctionDecl { name: "foo", params: [x], ... }` + implicit Return of last expr |
| Lua | `function foo(x) end` | `FunctionDecl { name: "foo", params: [x], ... }` |
| Fortran | `SUBROUTINE FOO(X)` | `FunctionDecl { name: "FOO", params: [X], is_sub: true, ... }` |
| Go | `func foo(x int) {}` | `FunctionDecl { name: "foo", params: [x], ... }` |

Lambdas/closures/arrow functions all normalize to `ExprKind::Lambda`. The compiler handles closures identically regardless of source language — captured variables use the shared WASM GC closure model (`emitter/closures.rs`: env array + `array.get`/`array.set`).

**Class method closures need shared env.** `compile_lambda_direct` creates shared env arrays for captured locals, but the class method compiler in `classes.rs` has its own path. Both paths now create shared envs when `current_closure_captured_locals` is non-empty. Destructuring bindings (`const {a,b} = x`) also sync to the shared env via `compile_destructure_bind`. Without these, closures inside class methods (e.g. `[Symbol.iterator]()` returning `{next(){...}}`) can't access the method's locals.

#### Iterators and spread — pure WASM bytecode (`emitter/generators.rs`)

All spread (`...x`), `for...of`, and `Array.from(iterable)` use the same iterator drain in `generators.rs`. The protocol follows ECMA-262 §7.4 and is implemented in pure WASM bytecode:

1. **§25.1.2 check**: if obj has own `next` property → it IS already an iterator. Use directly.
2. **§7.4.2 GetIterator**: `STRUCT_GET "Symbol(@@iterator)"` resolves through TypeRegistry for built-in types (Array→`values`, Map→`entries`, Set→`values`, String→`iterator`). Fallback to walker-normalized `"iterator"` key for custom JS classes.
3. **§7.4.5 IteratorStep loop**: `STRUCT_GET "next"` → `CALL_REF 1` → check `done` → push `value`. Pure WASM block/loop/br_if.
4. **§23.1.2.1 array-like fallback**: no `next` and no `[Symbol.iterator]` → read `length`, loop 0..length with `ARRAY_GET`. Pure WASM.

The host `iterForOf` (`ecma:object`) is a compatible fast path used by `emit_iter_for_of` — both produce identical results. The WASM path is lazy (true ECMA), the host path materializes eagerly.

Strings are opaque per `wasm:js-string` — `ecma:string.iterator` is a host function (same category as `charCodeAt`). This is the only host call in the iterator path.

#### Shared emitters — the cross-language backbone (`src/emitter/`)

The emitter modules are the shared infrastructure that makes "X over JS" work. Every language's profile routes operations through these modules. They emit standard WASM opcodes + `ecma:*` host calls — never custom opcodes, never language-specific host calls.

| Module | What it covers | How languages use it |
|---|---|---|
| `collections.rs` | Array ops: push, pop, slice, join, sort, reverse, indexOf, contains, concat, flat, map, filter | Profile: `emit = "common:collections.push"`. Every language's array.push / list.append / table.insert routes here. |
| `dict.rs` | Dict/map: new, set, get, has, delete, keys, values, entries | Profile: `emit = "common:dict.has"`. Python dict, Lua table, JS object literals route here. |
| `strings.rs` | String interpolation, concat, toString | Used directly by compiler for template literals, f-strings, `${}` interpolation across all languages. |
| `math.rs` | abs, floor, ceil, sqrt, round, min, max, pow, trig | Profile: `emit = "common:math.abs"`. All languages' math builtins route here. |
| `classes.rs` | Object construction, methods, inheritance, type stamping | Used by compiler for all class declarations across all languages. |
| `functions.rs` | Function chunks, default params, closures, async | Used by compiler for all function declarations. |
| `closures.rs` | Shared env closure model — WASM GC array for captured vars | Used by compiler for all closures/lambdas in all languages. |
| `errors.rs` | try/catch/finally, exception construction, throw | Used by compiler for all error handling across all languages. |
| `generators.rs` | Iterator protocol (ECMA §7.4), spread/for-of drain, generator stack-switching | `emit_drain_custom_iterable` (pure WASM iterator loop), `emit_drain_into_array` (generator drain via resume/suspend), `emit_spread_iterable` (dispatch: generator→stack-switch, iterable→WASM drain, array-like→index loop). |
| `promises.rs` | Promise chain (.then/.catch/.finally) via WASM JSPI | Used by compiler for promise chain method calls. Pure WASM — no host calls. |
| `loops.rs` | for-in iteration, HOF (map, filter, forEach, reduce) | Used by compiler for all loop constructs. |
| `io.rs` | Print via WASI, file I/O | Profile: `emit = "print"`. Every language's print/echo/write routes here. |
| `convert.rs` | Type conversions (toInt, toFloat, toString, toBool) | Profile: `emit = "common:to_int"`. All languages' type conversion builtins. |

**The pattern:** language grammar parses → walker normalizes to JS-shaped AST → profile maps builtin names to `common:<module>.<fn>` → emitter module emits standard WASM bytecode. Adding a new language means writing a grammar + walker + profile that maps to these existing emitters. You should almost never need to add a new emitter module.

### Step 4: Write the profile

Create `profile` (TOML). This tells the compiler how to emit bytecode for this language.

**Required `[info]` section** (for language registry):

```toml
[info]
name = "<lang>"
extensions = ["<ext1>", "<ext2>"]
```

**Required `[compiler]` settings:**

```toml
[compiler]
function_return = "explicit"       # "explicit" (return stmt), "result_slot" (Pascal/VB), "last_expression" (Ruby)
self_keyword = "this"              # "this" (JS/C#), "me" (VB), "self" (Python/Ruby/Pascal)
constructor_name = "constructor"   # "constructor" (JS), "New" (VB), "__init__" (Python), "Create" (Pascal)
case_sensitive = true              # false for VB, COBOL
string_indexing = "zero_based"     # "one_based" for VB, Pascal
```

**Optional `[compiler]` flags:**

```toml
base_keyword = "super"             # super/base/MyBase/inherited
implicit_self_fields = true        # bare field names → Me.field (VB, Pascal)
separated_methods = true           # methods defined outside class body (Pascal)
array_upper_bound_inclusive = true  # Dim arr(5) = 6 elements (VB)
parens_for_index = true            # arr(i) not arr[i] (VB)
entry_point = "main"               # auto-call if defined
partial_classes = true             # merge Partial Class (VB, C#)
with_block = true                  # With...End With (VB, Pascal)
hoist_var = true                   # var hoisting (JS)
dynamic_add = true                 # + overloaded for concat (JS)
explicit_self_param = true         # def method(self, x) (Python)
enum_as_ordinals = true            # enum values are integers (Pascal)
switch_fallthrough = true          # switch falls through without break (JS)
```

**`[array_methods]` section** — routes HOF methods (map, filter, sort) through the compiler's HOF dispatch instead of the generic method call path. Each entry maps a method name to a stdlib name that the compiler normalizes before dispatching:

```toml
[array_methods]
map       = "__array_map"
filter    = "__array_filter"
forEach   = "__array_forEach"
find      = "__array_find"
findIndex = "__array_findIndex"
includes  = "__array_includes"
reduce    = "__array_reduce"
sort      = "__array_sort"           # JS: 2-arg comparator sort
OrderBy   = "__array_sort_by_key"    # .NET LINQ: 1-arg key selector sort
some      = "__array_some"
every     = "__array_every"
```

The `sort` vs `sort_by_key` distinction: JS `arr.sort((a,b) => b-a)` passes a 2-arg comparator. .NET `list.OrderBy(x => x.name)` passes a 1-arg key selector. Both route through stdlib bytecode functions (`__stdlib_sort_with_comparator` and `__stdlib_sort_by_key` respectively).

**Emit formats for `[builtins]` and `[value_methods]`** — listed **in preference order**. Always pick the first format that satisfies the operation. WASM-first; language-specific Rust adapters last.

| # | Format | What it does | Example | When to use |
|---|---|---|---|---|
| 1 | `"opcode:<name>"` | Single WASM opcode | `"opcode:f64_abs"`, `"opcode:str_to_upper"` | Spec-conformant 1:1 mapping to a WASM opcode (math, string ops, comparisons). Cheapest path. |
| 2 | `"intrinsic:<name>"` | Multi-opcode template from `[intrinsics]` table | `"intrinsic:ubound"`, `"intrinsic:strtr"` | Small fixed sequence of opcodes; no host call. Compile-time constant data (`Math.PI = 3.14...`) goes here. |
| 3 | `"common:<module>.<fn>"` | `emitter::dispatch::emit_common` | `"common:dict.has"`, `"common:collections.push"`, `"common:strings.to_upper"` | Operation has a language-agnostic shape — collections, dict ops, threading, iteration helpers. Same dispatch every language uses. |
| 4 | `"common:<lang>.<fn>"` | Language-runtime adapter under `emitter/<lang>/` | `"common:php.date"`, `"common:php.array_chunk"`, `"common:dotnet.linq_first"`, `"common:dotnet.console_writeline"` | Function semantics are .NET-shaped or language-specific (PHP `date()` format codes, .NET `Console.WriteLine` boolean capitalization, LINQ `Aggregate(seed, fn)` arg order). The adapter inline-emits opcodes composing existing host fns. **This is also the home for .NET BCL surfaces (Console, StringBuilder, DateTime, LINQ, Streams, …) shared across VB and C#.** |
| 5 | `"host:<module>:<func>"` | `call_import(module, func)` | `"host:ecma:date:now"`, `"host:wasi:cli:log"` | Direct call into ECMA-262 / WASI / Component-Model spec surface. All host fns are ecma-shaped or wasi-shaped — the ONLY exception is `vybe:gui` for GUI. Use the spec name (`ecma:date.parse`, `wasi:filesystem.readFile`). NEVER invent module names. NEVER add language-specific host namespaces. |
| 6 | `"stdlib:<name>"` | Generic shared bytecode chunk via `__vybe_<name>` global | `"stdlib:sorted"`, `"stdlib:range"`, `"stdlib:tostring"` | The function is shared across multiple languages and benefits from amortizing chunk overhead (sort with comparator, range generator, structuredClone). NEVER for language-specific fns — those go through (4). |
| 7 | `"invoke:<MethodName>"` | Polymorphic dispatch via `wasm:js-value.invokeMethod` | `"invoke:Peek"`, `"invoke:ContainsValue"`, `"invoke:TryGetValue"` | The receiver type is unknown at compile time and the method exists on multiple shapes (Stack vs Queue both have `Peek`, but they read different ends). The runtime walks the prototype/type-registry chain to dispatch. Cheaper than introducing per-shape adapters when the call surface is small. |
| 8 | `"mutate:<op>"` | In-place `var = var OP arg` | `"mutate:add"` (Inc/`++`), `"mutate:sub"` (Dec/`--`) | Self-modifying statements. |
| 9 | `"print"` / `"str_length"` / etc. | Bespoke helper from `emitter::*` | `Console.WriteLine`, `Len()` | Legacy convenience names; all forward to one of the above. |
| 10 | `"noop"` | Emit nothing | Layout no-ops, `randomize`, `free`, PHP `error_reporting` | Stub for runtime configuration calls that don't affect compiled bytecode. |

**Layer priority (the user-mandated rule — non-negotiable):**

1. **WASM opcodes** — if a WASM opcode already does what you need, use it directly. This is the fastest and most correct path.
2. **Compiler emitters** — `emitter/<feature>.rs` or `languages/<lang>/src/emitter/<feature>_adapter.rs`. Compose existing opcodes and host fns into bytecode sequences. No new host fns needed.
3. **ecma:\* host functions** — if an existing `ecma:*` function maps directly to what you need, use `self.import("ecma:<module>", "<name>")`. ecma:\* is ECMA-262/402/WebIDL only.
4. **Language-specific adapters** — `languages/<lang>/src/emitter/<feature>_adapter.rs` for runtime semantics that can't be walker-normalized. Emits bytecode composing existing fns.

**What is NOT an option — ever:**
- Adding custom (non-WASM) opcodes to `vybe_runtime/` — NEVER, under any circumstances. All past custom opcodes have been removed.
- Modifying the VM — it is LAST RESORT, only for fixing bugs in existing WASM opcode handling. Even WASM-compliant changes require explicit per-change user approval.
- Adding host functions to `platforms/{ecma,wasi,web,node,vybe}` — requires explicit per-function user approval. Host fns are ALWAYS ecma-shaped (`ecma:*`) or wasi-shaped (`wasi:*`). The ONLY non-spec namespace is `vybe:gui` for GUI. No language-specific host namespaces (`php:*`, `python:*`, `vb:*`). The default path is the emitter/adapter pattern.
- JS polyfills, `__vybe_*` helpers, or any `.js` runtime files.

Walker normalization is orthogonal — it produces a JS-shape common AST regardless of which layer the compile path picks. Behavior must remain language-faithful (PHP rounding semantics differ from JS).

**Forbidden bindings** — these pattern-fail every code review:
- `host:php:*`, `host:vb:*`, `host:python:*` — `php`/`vb`/`python` aren't host namespaces. ECMA-262 + WASI cover what they cover; everything else lives in **adapters** (path 4) or **walker rewrites**.
- `host:ecma:date:phpDate` — `ecma:*` is the ECMA-262 / 402 / WebIDL spec surface. Language-specific helpers don't go here. Build a `common:php.date` adapter instead.
- `stdlib:php_<anything>` for new code — the `stdlib:` prefix is for SHARED helpers; language-specific helpers route through `common:<lang>.<name>` adapters.

### Step 5: Write lib.rs (the crate root)

Each language is its own crate `vybe_language_<lang>`. Its `src/lib.rs` is the crate root:
the pest derive struct, `parse()`, `profile_source()`, and the registration entry points
(`register()` + `struct Plugin`).

```rust
pub mod walker;
pub mod normalize_class;
pub mod emitter;            // if the language has common:<lang>.* adapters

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct <Lang>Parser;

/// Parse source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (also the dylib entry point).
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "<lang>",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch), // or None
        normalize_class: Some(normalize_class::normalize_class), // or None
        register_tree: None, // Some(tree_register::register) if it mounts a namespace tree
    });
    // Only if the shared compiler must call back into this language:
    // vybe_runtime::registry::register_hooks("<lang>", vybe_runtime::registry::LanguageHooks {
    //     value_eq: Some(...), proxy_get: Some(...), parse_eval: Some(...), ..Default::default()
    // });
}

/// This crate as a `vybe_runtime::Plugin` — `init` registers everything it provides.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str { "<lang>" }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) { register(); }
}
```

### Step 6: Register the language

Vybe has a layered registration system. Each layer registers different things, and they compose automatically. You should never need to edit shared dispatch code, CLI code, or host registration when adding a language.

#### 6a. File extensions — registered in the profile

Extensions are declared in the profile's `[info]` section. The runtime reads them via `languages::find_by_extension()` — no Rust code changes needed.

```toml
[info]
name = "pascal"
extensions = ["pas", "pp", "dpr", "lpr"]
```

Current extension registry (from profiles):
| Language | Extensions |
|---|---|
| JS | `.js`, `.mjs` |
| VB | `.vb` |
| PHP | `.php`, `.phtml` |
| Python | `.py`, `.pyw` |
| C# | `.cs` |
| C | `.c`, `.h` |
| Pascal | `.pas`, `.pp`, `.dpr`, `.lpr` |
| Dart | `.dart` |
| Ruby | `.rb` |
| COBOL | `.cob`, `.cbl`, `.cobol` |
| Fortran | `.f90`, `.f95`, `.f03`, `.f08`, `.f18`, `.f` |
| Go | `.go` |
| Lua | `.lua` |
| Java | `.java` |

#### 6b. Project files — registered in `projects/mod.rs`

Multi-file projects are dispatched by project file extension in `crates/vybe_compiler/src/projects/mod.rs::load()`:

```rust
match ext.as_str() {
    "vybe"                          => vybe::load(path),
    "vbproj"                        => vbproj::load(path),
    "csproj" | "pyproj" | "ipyproj" => managed_msbuild::load(path),
    _                               => single_file::load(path, &ext),
}
```

| Extension | Loader | Languages |
|---|---|---|
| `.vybe` | `projects/vybe.rs` | Multi-language manifest |
| `.vbproj` | `projects/vbproj.rs` | VB.NET (MSBuild XML, `<Compile>` items, form designer files) |
| `.csproj` | `projects/managed_msbuild.rs` | C# (MSBuild XML, `<Compile>` items) |
| `.pyproj`, `.ipyproj` | `projects/managed_msbuild.rs` | Python (IronPython MSBuild format) |
| any source ext | `projects/single_file.rs` | Single-file compile (auto-detects language from extension) |

All loaders produce a `Bundle { name, language, sources, wasm_files, entry_point }` — the universal compile unit. The CLI and tests consume `Bundle` the same way regardless of whether it came from a single `.js` file or a multi-file `.vbproj`.

#### 6c. Language registration — the plugin registry (NOT a central `all()` list)

Registration is **inverted**: there is no hand-maintained plugin array anywhere in the tree.
Each plugin crate ends with `vybe_runtime::register_plugin!(Plugin);`, which submits it at
LINK time; a `build.rs` generates one `extern crate` per plugin dependency from the crate's own
Cargo.toml so the rlib is actually linked. Adding a plugin is a **Cargo edit only** — one
registry, one init loop (`vybe_runtime::init_all_registered`), no list in code.
Each language crate registers *itself* by calling `register()` (Step 5), which pushes a
`LanguageDef` descriptor into `vybe_runtime::registry`. The shared compiler only ever
looks languages up **by name** through that registry
(`crate::languages::find_by_name`, `emit_dispatch_for`, `registry::hooks(name)`), so it
never names a concrete language crate and a language can live in (eventually load as) its
own dylib.

```rust
// languages/<lang>/src/lib.rs::register()  (see Step 5 for the full form)
vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
    name: "<lang>",
    parse,                                           // source → common AST
    profile_source,                                  // embedded TOML
    emit_dispatch: Some(emitter::dispatch::dispatch), // routes common:<lang>.* (or None)
    normalize_class: Some(normalize_class::normalize_class), // class → NormalClass (or None)
    register_tree: None,                             // Some(...) to mount a namespace tree
});
```

**Who calls `register()`?** Whoever assembles a compile/run: the `vybex` CLI registers all
languages at startup; a language's **test helper** calls its own `register()` once (guarded
by `Once`) before compiling. `crate::languages::mod.rs` in `vybe_compiler` is now a *thin
facade* over `vybe_runtime::registry` (`all()`, `find_by_name`, `find_by_extension`,
`emit_dispatch_for`) — you do **not** edit it to add a language.

**`emit_dispatch`** is `Some(...)` if the language has its own `common:<lang>.*` builtins
with a `languages/<lang>/src/emitter/dispatch.rs`; `None` if it only uses shared common ops
(`collections`, `strings`, …) and/or a platform like `dotnet`.

**`LanguageHooks`** (via `register_hooks("<lang>", …)`) is the small set of callbacks the
shared compiler invokes *back* into a language — `value_eq` (Python), `relational_compare`/
`str_getcsv`/`normalize_source` (PHP), `proxy_*` (JS), `parse_eval` (JS eval-string parse).
All fields are `Option` and default to `None`; register only what you have.

#### 6d. Emit dispatch registration — how `common:*` routes work

When the profile says `emit = "common:php.date"`, the dispatch chain is:

```
Profile: emit = "common:php.date"
    ↓
emitter/dispatch.rs::emit_common("php.date")
    ↓  splits on "." → prefix="php", op="date"
languages::emit_dispatch_for("php")
    ↓  looks up LanguageDef::emit_dispatch in the registry
languages/php/src/emitter/dispatch.rs::dispatch("date", chunks, current, argc, line)
    ↓  matches "date"
languages/php/src/emitter/datetime_adapter.rs::emit_date(chunks, current, argc, line)
```

Three dispatch levels:
1. **Shared** (`collections`, `dict`, `strings`, `threading`, `channels`, `object`) — handled directly by `emitter/dispatch.rs`
2. **Language-owned** (`php`, `vb`, `js`, `python`, `dart`, `ruby`, `cobol`, `fortran`, `go`) — delegated to `languages/<lang>/src/emitter/dispatch.rs` via the `LanguageDef::emit_dispatch` field
3. **Platform-owned** (`dotnet`) — delegated to `platforms/dotnet/src/emitter/dispatch.rs` via `emitter::platform_emit_dispatch()`

Adding a language's emit dispatch:
1. Create `languages/<lang>/src/emitter/mod.rs` with `pub mod dispatch;`
2. Create `languages/<lang>/src/emitter/dispatch.rs` with `pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool`
3. Add `pub mod emitter;` to `languages/<lang>/mod.rs`
4. Set `emit_dispatch: Some(<lang>::emitter::dispatch::dispatch)` in the `all()` entry
5. **Zero edits to `emitter/dispatch.rs`** — the central dispatcher resolves your prefix through the registry

#### 6e. Host function registration — the platform crates

Host functions live in `platforms/{ecma,wasi,web,node,vybe}` and are **OFF-LIMITS without
explicit user authorization**. Each platform is a `vybe_runtime::Plugin` that registers its
own host fns in `init`. Registration happens at VM startup through the ONE loop, which walks
the plugin registry — no platform is named:

```
vybe_compiler::primitives::platforms::register_platforms_all(&mut vm)
    ↓
vybe_runtime::init_all_registered(vm, caps)   ← THE loop, over THE registry
    ↓
  phase 1 — every linked plugin's `Plugin::init(fw)`, skipped when its
            `required_capability()` is not granted:
    platforms/ecma  → ecma:array, ecma:string, ecma:object, ecma:math, …
                      + wasm:js-string (real WASM proposal)
    platforms/web   → web:crypto, web:url, web:encoding, …
    platforms/wasi  → wasi:cli, wasi:clocks, wasi:random, wasi:filesystem,
                      wasi:io/streams, wasi:http, wasi:crypto, wasi:sql,
                      wasi:sockets   (each capability-gated internally)
    platforms/node  → node:fs, node:os, node:path, node:process, node:http, …
    platforms/vybe  → vybe:gui (drawing always; widgets under feature `gui`
                      when Gui is granted; otherwise it installs its own
                      no-op stubs)
    ↓
  phase 2 — every plugin's `Plugin::finalize(fw)`, once all host fns exist:
    platforms/ecma  → globalThis + constructor↔prototype wiring (resolves
                      host fns other plugins registered, by index)
    platforms/vybe  → TypeRegistry vtables / control types
```

The two phases are separate calls (`init_registered_plugins` / `finalize_registered_plugins`)
so a test harness can **override a host fn between them** — `register_host_fn` appends a new
index rather than replacing, and `finalize` resolves by index, so an override registered after
`finalize` would not be seen by the globals. Harnesses that capture output do:
`init_platforms(vm)` → override host fns → `finalize_platforms(vm)`.

Host fn namespaces:
| Namespace | Spec | Examples |
|---|---|---|
| `ecma:*` | ECMA-262 / 402 / WebIDL | `ecma:array`, `ecma:string`, `ecma:math`, `ecma:date`, `ecma:promise` |
| `wasi:*` | WASI proposals | `wasi:filesystem`, `wasi:cli`, `wasi:clocks`, `wasi:http`, `wasi:io` |
| `wasm:*` | WASM CG proposals | `wasm:js-string` (js-string-builtins) |
| `web:*` | WHATWG / W3C | `web:crypto`, `web:url`, `web:encoding` |
| `node:*` | Node.js builtins | `node:fs`, `node:os`, `node:path`, `node:process` |
| `vybe:gui` | Vybe-specific (ONLY exception) | GUI controls, canvas, dialogs |

**Capability-based security:** host fns are grouped by capability (`Console`, `Clock`, `Random`, `FileRead`, `FileWrite`, `Http`, `Sockets`, `Database`, `Crypto`, `Process`, `Environment`, `HttpServer`). The `--sandbox` CLI flag uses `Capabilities::safe()` which disables filesystem, network, process, etc.

#### 6e-bis. Declared host functions — `HostFnDecl` and the arity check

`register_host_fn(module, name, closure)` registers a function that says **nothing
about itself**: not how many arguments it takes, not what it returns, not whether a
handle it receives is owned or borrowed. A caller that passes three arguments to a
four-argument function gets a silent `Value::Null` in the missing slot and fails
somewhere else entirely.

`VM::register_host(HostFnDecl)` is the same registration plus a Component Model
signature:

```rust
vm.register_host(
    HostFnDecl::new("web:dom", "appendChild", Box::new(move |_ctx, args| { … }))
        .with_sig(FuncSig {
            name: "append-child".into(),
            params: vec![ValType::Borrow("document".into()),
                         ValType::Borrow("node".into()),
                         ValType::Borrow("node".into())],
            results: vec![ValType::Borrow("node".into())],
        })
        .method_on("document"),
);
```

Declaring is **per-function and optional**. An undeclared registration behaves
exactly as it always did — `declared_host_arity` answers `None`, which means
UNKNOWN, never zero. There is no flag day and no requirement to convert a module
all at once.

What a declaration buys, at COMPILE time: `Chunk::emit_call` — the one point every
host call funnels through, the 282 `emit_host_call` sites and the 181 that reach for
`emit_call` directly — checks the emitted argc against the declaration and prints

```
[vybe] host call arity: web:html:setValue declares 3 parameter(s), called with 2 (line 88)
```

A warning is a **finding**, not noise. It means either the declaration is wrong or
the call site is, and both need reading. Do not "fix" it by trimming the declaration
to match a wrong caller.

Three conventions worth knowing before you write one:

- **`results` is the Component Model result, not the VM's stack convention.** A void
  operation declares `results: vec![]` even though every host call leaves one value
  on the stack (which is why the `gui::emit_*` helpers drop it). A setter declaring
  `vec![]` while its closure answers `Value::Null` is those two conventions meeting,
  not a mismatch.
- **`Own(T)` vs `Borrow(T)`.** `createElement` returns `own<node>` — a fresh handle
  the caller is responsible for. `appendChild` takes and returns `borrow<node>`: it
  neither consumes its parent nor its child. Without the distinction every handle
  looks owned and a host that dropped one would be indistinguishable from one that
  did not.
- **Nullability is part of the signature.** An absent DOM attribute is
  `option<string>` (`getAttribute` answers null); an unset CSS property is plain
  `""` (`getStyleProperty`). Those genuinely differ, and a declaration is the only
  place that difference is written down.

**The Component Model has no optional parameter.** A WebIDL `optional` is `option<T>`
and still positional. If a closure reads a trailing argument only "when present",
either the call sites pass it explicitly or the declaration states the arity they
actually use — declaring the larger arity fires the check on every real call, which
is the failure the mechanism exists to prevent rather than cause.

**Some functions should stay undeclared, and that is not laziness.** `FuncSig`
carries a fixed `Vec<ValType>`, so a genuinely variadic or optional-arg function
has no honest single arity. `setTimeout(handler, optional timeout, any... args)`
is the worked example: `setTimeout(fn)` is ordinary guest code, so declaring `2`
would warn on correct programs — and a warning that fires on correct code teaches
everyone to ignore the one that matters. `console.log` is the same. Leave them
undeclared; `None` means UNKNOWN, which is true, and declaring is per-function
precisely so a module can be half declared.

**Where the check pays, and where it mostly does not.** It catches **hand-built**
argc — `emit_host_call(idx, 3)` in `primitives/gui.rs` and the language adapters,
where a dropped operand silently becomes `Value::Null`. For an ordinary `ecma:*`
call the argc comes from the guest program's own argument list, so a "mismatch"
is usually the guest legitimately calling a variadic. Declare where emitters
build the call, not everywhere.

When several functions in one module share a shape, register them through a small
local helper so the signature and the closure cannot drift apart — see
`dom_fn`/`html_fn`/`css_fn` in `platforms/web/src/html.rs`, where all 38 functions go
through three helpers and `register_host_fn` no longer appears at all.

Adding a host fn still **requires explicit per-function user approval** (level 9 in
the fix ladder). Declaring one that already exists does not — it adds metadata and
changes no behaviour.

**Adding a new language NEVER touches `platforms/{ecma,wasi,web,node,vybe}`.** The host surface is spec-shaped and shared by all languages. Language-specific runtime semantics go in emitter adapters.

#### 6f. No cross-language emitter re-exports

Now that each language is its own crate, there are **no** `pub use crate::languages::<lang>::emitter`
aliases in `vybe_compiler`. A language's emitter code is reached only through its
registered `emit_dispatch` fn-pointer (by name, via the registry). Compiler code never
writes `crate::emitter::php::…`; if it needs php behavior it goes through
`emit_dispatch_for("php")`.

### Step 7: Tests live with the crate

Tests belong to the language crate: `languages/<lang>/tests/<lang>/`. Match the per-file
`#[test]` count exactly — porting/adding is not a chance to drop tests.

Create under `languages/<lang>/tests/<lang>/`:
- `main.rs` — `mod` declarations for `helpers` + all test files
- `helpers.rs` — `run_<lang>()` plus optional GUI variants

The crate's `Cargo.toml` declares the test target and dev-deps:
```toml
[dev-dependencies]
vybe_compiler = { workspace = true }   # the compiler AND the shared emitter
vybe_runtime = { workspace = true }
vybe_ast      = { workspace = true }
# Platform crates this language needs at RUNTIME must be listed so they are
# linked — that is how they reach the plugin registry (see build.rs below):
#   vybe_platform_vybe / _ecma / _wasi / _node / _web  (the runtime host set,
#   already pulled in transitively by vybe_compiler), plus the language's own
#   platform: dart→_flutter, pascal→_plib, csharp/vb→_dotnet, c→_libc.
# + vybe_platform_wasm if a helper calls write_wasm; + other language crates for
#   cross-language tests; + uuid/chrono if used, etc.

[[test]]
name = "<lang>"
path = "tests/<lang>/main.rs"
```

The helper registers the language once, then compiles + runs:

```rust
use std::sync::{Arc, Mutex};
use vybe_runtime::{VM, Value, HostContext};

use std::sync::Once;
static REGISTER: Once = Once::new();

/// Run source through the vybe_compiler pipeline; return console output.
pub fn run_<lang>(src: &str) -> Vec<String> {
    // Register THIS language crate with the plugin registry exactly once.
    REGISTER.call_once(vybe_language_<lang>::register);

    let module = vybe_language_<lang>::parse(src).expect("parse failed");
    let profile = vybe_compiler::profile::parse_profile(vybe_language_<lang>::profile_source())
        .expect("Failed to parse profile");
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module).expect("compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("run failed");
    let result = output.lock().unwrap().clone();
    result
}
```

If your language has GUI host fns, add a sibling `run_<lang>_gui()` that uses `vybe_platform_vybe::init_platforms_with_gui` (then `finalize_platforms`) and returns `(VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>)`. See `languages/vb/tests/vb/helpers.rs` for the canonical shape (it also includes `run_vb_gui_capture_msgbox` for tests that need to assert on dialog calls — see "MsgBox capture in tests" below).

For a helper that must run programs which `include`/`eval` **at runtime** (php/python/js),
use the dynamic-compile service instead of a one-shot compile:
```rust
let mut runtime = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
runtime.compile_and_run_source(src, language, virtual_path).expect("run failed");
```
(`RuntimeCompilerService` lives in `vybe_compiler`, not `vybex` — see "Dynamic compilation
& eval" in `fix_tests.md`.)

#### MsgBox capture in tests

Production `vybe:gui::msgBox` shows a real native dialog inline via `vybe_widgets::dialogs::MessageBox::info`. There is **no queue** to inspect. Tests that need to assert on msgbox calls register a per-test override of the host fn AFTER `init_platforms_with_gui` runs (and before `finalize_platforms`), swapping the production impl for one that pushes onto a captured `Vec<(String, String)>`. This is the same pattern tests use to capture `wasi:cli::log` output. Don't add a queue. Don't add a `pending_*` field on `GuiState`. Override the host fn.

### Step 8: Test

```bash
echo '<hello world code>' > /tmp/test.<ext>
cargo run -p vybex --bin vybex -- /tmp/test.<ext>
cargo test -p vybe_language_<lang> --test <lang>
```

---

## The shared emitter — MUST use for ALL compilation

The shared cross-language emit helpers live in **`crates/vybe_compiler/src/primitives/`**,
alongside the AST walkers that call them. There is no separate emitter crate and no
separate `emitter/` module: **`vybe_compiler::primitives` IS the emitter.** Inside the
compiler it is imported as:

```rust
use crate::primitives as common;
```

From outside (a language or platform crate) the same helpers are reached as
`vybe_compiler::primitives::ops`, `…::collections`, `…::strings`, and so on.

Where a topic has both halves — `expressions`, `events`, `reflection` — the `impl Compiler`
walkers and the free `&mut Chunk` primitives live in the **same file**.

The compiler MUST use these helpers (i.e. `common::*`) for all bytecode emission. This ensures:
- Cross-language interop (JS class inherits from VB class)
- WASM compliance (standard opcodes, proper import table)
- Consistent behavior across all languages

Call emitter functions on `&mut self.chunks[0]` for functions that add imports (WASM import table is module-level on chunk 0). For functions that only emit opcodes (abs, floor, etc.), use `self.chunk()`.

| Module | What it does | Key functions |
|---|---|---|
| `classes` | Object construction, fields, methods, property accessors, inheritance, type stamping, cross-language aliases, static methods, getters/setters | `emit_new_typed_object`, `emit_bind_method_with_aliases`, `emit_save_base_method`, `emit_store_super`, `emit_attach_static_method`, `emit_bind_getter`, `emit_bind_setter`, `emit_constructor_return`, `emit_store_constructor`, `emit_super_call_store_result`, `emit_instanceof_chain`, `emit_auto_init_component`, `register_type` |
| `functions` | Function chunks, default params, closures, async/await, spread args | `create_function_chunk`, `emit_default_param_start/end`, `emit_ref_func`, `emit_await`, `emit_async_wrapper` |
| `collections` | Array operations, sorting, contains, indexOf | `emit_array_new`, `emit_get`, `emit_len`, `emit_push`, `emit_pop`, `emit_slice`, `emit_join`, `emit_reverse`, `emit_contains`, `emit_index_of`, `emit_sorted`, `emit_concat` |
| `dict` | Dictionary/map/set operations | `emit_new`, `emit_set`, `emit_get`, `emit_method_has`, `emit_method_clear`, `emit_method_size`, `emit_keys`, `emit_values` |
| `loops` | For-in iteration, higher-order (map, filter, forEach, reduce, some/every) | `emit_for_in_start/end`, `emit_map`, `emit_filter`, `emit_foreach`, `emit_reduce`, `emit_any_every`, `emit_loop_start/cond/end` |
| `errors` | Try/catch/finally, exception construction, cross-language exception normalization | `emit_try_start`, `emit_try_end`, `patch_catch`, `emit_throw`, `emit_exception_new_finalize`, `is_exception_type`, `canonical_exception_name` |
| `expressions` | Ternary, short-circuit AND/OR, null coalescing | `emit_ternary_start/middle/end`, `emit_and_start`, `emit_or_start` |
| `strings` | String operations as WASM opcodes | `emit_length`, `emit_to_upper`, `emit_to_lower`, `emit_trim`, `emit_substring`, `emit_index_of`, `emit_replace`, `emit_split`, `emit_concat`, `emit_repeat` |
| `math` | Math — WASM opcodes + host fallbacks | `emit_abs`, `emit_floor`, `emit_ceil`, `emit_sqrt`, `emit_round`, `emit_min`, `emit_max`, `emit_pow`, `emit_sin`, `emit_cos`, `emit_log`, `emit_random` |
| `io` | Print/input via WASI, file I/O | `emit_print`, `emit_print_with_import`, `emit_input`, `emit_read_file`, `emit_write_file`, `emit_open_file`, `emit_close_file`, `emit_print_file`, `emit_input_file`, `emit_line_input` |
| `convert` | Type conversions | `emit_to_int`, `emit_to_float`, `emit_to_string`, `emit_to_bool` |
| `threading` | Atomics, WASM stack switching | `emit_task_run`, `emit_thread_new`, `emit_thread_join`, `emit_sleep` |
| `gui` | Canonical GUI host fn naming + emit helpers — single source of truth for what a GUI control IS, what it's called on the host side, and how to emit creating/binding/configuring it. Every framework wrapper goes through here. | `canonical_control_name`, `host_fn_new_control`, `HOST_FN_SET_PROPERTY`, `HOST_FN_BIND_EVENT`, `HOST_FN_RUN_APPLICATION`, `emit_new_control`, `emit_bind_event`, `emit_add_child`, `emit_run_application`, `emit_get_control_name` |
| `invoke` | Dynamic method invocation for polymorphic receivers. When the receiver type is unknown at compile time, routes `receiver.method(args)` through `wasm:js-value.invokeMethod` — on v8 this resolves natively via the prototype chain, on Vybe VM the host handler does the same walk. Use for dynamically-typed method calls; use the typed `collections::emit_*` / `strings::emit_*` helpers for statically-typed receivers. | `emit_invoke_method` |
| `components` | Component Model metadata for cross-language linking. After compiling to chunks, build a `Component` with proper export declarations so the Linker can resolve cross-language imports (Python defines `greet`, Dart calls `greet`). | `build_component` |
| `canonical` | Canonical method/exception/event naming for cross-language interop. Dunder-style canonical names (`__len__`, `__str__`, `__upper__`, `__contains__`, …) unify Python `len()`, JS `.length`, C# `.Length`, VB `Length()` at one call site. | `cross_language_aliases`, `canonical_exception_name` |
| `dispatch` | `common:<name>` dispatcher used by profiles via `emit = "common:..."`. Owns the shared `common:<cat>.*` keys directly and **delegates language/platform prefixes through the registry** (`languages::emit_dispatch_for`) — each language's arms live in `languages/<lang>/src/emitter/dispatch.rs`, `dotnet`'s in `platforms/dotnet/src/emitter/dispatch.rs`. No per-language code here. | `emit_common`, `emit_common_with_imports` |
| `type_registry` | Compile-time type table — `CompileTimeTypes`, `__tid_<name>` global emission, type-id stamping at construction | `register_type`, `lookup_type`, `emit_set_type_id` |
| `target` | The single `Target` enum (`Wasm`, `Native`, …) used by emit modules that need to know the deployment target | `Target` |
| `imports` | `IMPORT_ALIASES` table — maps `__vybe_*` global names to native host fns; used by stdlib polyfill override | `IMPORT_ALIASES` |
| `bundle` | Appends stdlib chunks and emits the preamble that installs `__vybe_*` globals. The preamble is polyfill-safe: if Vybe VM already populated the globals with optimized native fns, the preamble leaves them alone. On non-Vybe runtimes the bundled bytecode takes effect. | `emit_stdlib_preamble`, `finalize_with_stdlib` |
| `stdlib` | Pure WASM bytecode functions bundled into every compilation. Each function is addressable as `__stdlib_<name>` (chunk name) and `__vybe_<name>` (global call site). **Sort:** `sorted`, `sort_in_place`, `sort_with_comparator` (2-arg JS-style), `sort_by_key` (1-arg LINQ-style). **Sequence:** `range`, `reversed`, `enumerate`, `zip`, `splice`, `slice`, `slice_step`, `from` (arrayFrom). **Aggregation:** `sum`, `min`, `max`, `count`. **Math:** `pow`, `sin`, `cos`, `tan`, `asin`, `acos`, `atan`, `atan2`, `log`, `log10`, `exp`, `sinh`, `cosh`, `tanh`, `sign`, `clamp`, `floor`, `fmod`. **String:** `tostring`, `string_is_null_or_empty`, `string_is_null_or_whitespace`, `str_insert`, `str_remove_start`, `str_remove_range`, `string_raw`, `isnumeric`, `concat`. **Dict/Object:** `keys`, `hasproperty`, `assign`, `instanceof`, `deleteproperty`. **Array mutation:** `array_insert`, `array_remove_at`, `array_remove_value`, `array_insert_range`, `array_set_range`, `array_binary_search`, `array_reverse_range`, `array_last_index_of`. **VB/Pascal helpers:** `redim`, `dynmul`. | `build_stdlib`, `StdLib` |
| `dotnet` | .NET BCL/FCL frontend. Partitioned into `core` (language-agnostic BCL) and `winforms` (GUI framework adapter). See submodule breakdown below. | — |
| `dotnet::core` | Shared .NET metadata + adapters reusable by any .NET-shaped compiler (VB, C#, F#, …). Reference files: `imports`, `namespaces`, `host_map`, `types`, `component_classes`. **Adapters** (each is a `pub fn emit_<name>(chunks, current, line)` — or `(…, argc, line)` for variadic — that composes existing opcodes / host fns; routed via `dispatch.rs`): `console_adapter` (`Console.Write/WriteLine` with .NET-style bool capitalization + null→""), `linq_adapter` (`First`/`Last`/`Skip`/`Take`/`Average`/`FirstOrDefault`/`Distinct`/`Aggregate`/`OrderByDescending` + identity for `ToList`/`ToArray`), `array_adapter` (`Array.Clear`/`Copy`/`Resize`/`Sort`), `datetime_adapter`, `timespan_adapter`, `stringbuilder_adapter`, `string_format_adapter`, `format_picture_adapter`, `stream_io_adapter`, `process_adapter`, `sockets_adapter`. Add a new BCL surface here, NOT in language walkers — VB and C# pick it up by listing the `common:dotnet.<name>` emit target in their profile's `[value_methods]` / `[builtins]`. | `default_interface_imports`, `namespace_to_host_module`, `map_host_func`, `is_namespace_root`, `known_type_mappings`, `dotnet_core_component_descriptor`, `linq_adapter::emit_linq_*`, `console_adapter::emit_console_writeline`, `array_adapter::emit_array_*`, `datetime_adapter::*` |
| `dotnet::winforms` | WinForms GUI framework adapter. Has its own `imports`, `namespaces`, `host_map`, `types`, `component_classes` — same structure as `core` but for `System.Windows.Forms.*`. Includes the wrapper `classes/` submodule (see below). | `default_interface_imports`, `namespace_to_host_module`, `map_host_func`, `is_namespace_root`, `capitalize_control_name`, `is_noop_method`, `dotnet_winforms_component_descriptor` |
| `dotnet::winforms::classes` | The .NET WinForms **wrapper class hierarchy** for VB/C#. Real `Object → MarshalByRefObject → Component → Control → Form/Button/...` chain with property setters and methods compiled to bytecode. Per-file: `object`, `control`, `form`, `buttons` (ButtonBase/Button/CheckBox/RadioButton), `text` (TextBoxBase/TextBox/RichTextBox), `labels` (Label/LinkLabel), `lists` (ComboBox/ListBox/ListView/TreeView), `containers` (Panel/GroupBox/TabControl/TabPage), `progress` (ProgressBar/TrackBar/NumericUpDown), `dates` (DateTimePicker/MonthCalendar), `media` (PictureBox/WebBrowser), `grids` (DataGridView), `strips` (MenuStrip/ToolStrip/StatusStrip/ContextMenuStrip), `nonvisual` (Timer/BindingSource/ImageList/ToolTip), `drawing` (Graphics/Pen/SolidBrush), `dialogs`. New languages with their OWN class hierarchies get their own sibling submodule, NOT this one. | `DotnetClass`, `DotnetMethod`, `MethodTarget::{Host,DotnetCtor,Body}`, `MethodOp` DSL, `builder::{build_setter_chunk, build_method_thunk_chunk, build_constructor_chunk}` |
| `dotnet::resolver` | Dotted-name resolution algorithm shared by both `core` and `winforms`. Resolves `System.Threading.Thread.Sleep` → `(wasi:clocks, sleep)`. | `resolve_dotted_name`, `DottedResolution`, `ResolutionContext`, `resolve_interface_call` |
| `dotnet::class_exports` | Aggregates `DotnetClassExport` entries from `core::component_classes` and `winforms::component_classes` into a single flat list. `descriptor.rs` uses this to build `ComponentDescriptor` objects partitioned by `core` vs `winforms`. | `dotnet_class_exports`, `DotnetClassExport` |

## Framework wrappers, GUI, canvas, dialogs

Three layers, each independently shippable. PHP, Ruby, Python and other "non-GUI" languages can ignore all three. Languages that want to expose framework-shaped GUI APIs slot a wrapper into Layer 3.

```
┌─────────────────────────────────────────────────────────────────┐
│  Layer 3 — language-shaped façades                              │
│  emitter::dotnet::winforms::classes (BCL/FCL wrapper for VB/C#) │
│   - Form, Button, Label, Graphics, Pen, SolidBrush, …           │
│   - Each method: a MethodTarget::Body sequence of MethodOps     │
│     that emits real bytecode calling vybe:gui::canvas* host fns │
│  (Future: emitter::flutter::canvas, emitter::web::canvas2d      │
│  for Dart/JS canvas-shaped APIs.)                               │
└────────────────────────┬────────────────────────────────────────┘
                         │
┌────────────────────────▼────────────────────────────────────────┐
│  Layer 2 — platforms/vybe: canvas + gui host fns                │
│  - vybe:gui::canvas* host fns (HTML5-shaped: setFillColor,      │
│    moveTo, lineTo, fill, stroke, fillText, …)                   │
│  - vybe:gui::msgBox calls vybe_widgets::dialogs inline          │
│  - vybe:gui::new_<Type> + controlSetProperty + onEvent          │
│  Thin transport layer. Zero drawing/dialog logic.               │
└────────────────────────┬────────────────────────────────────────┘
                         │
┌────────────────────────▼────────────────────────────────────────┐
│  Layer 1 — vybe_widgets (the toolkit)                           │
│  - vybe_widgets::canvas::{Canvas trait, RecordingCanvas,        │
│    TinySkiaCanvas, Color, Font, Image, …}                       │
│  - vybe_widgets::canvas_widget::Canvas (a PanelWidget)          │
│  - vybe_widgets::dialogs::{MessageBox, FileDialog, FolderDialog}│
│  - All Form/Button/Label/etc. widgets                           │
│  Pure Rust toolkit, no VM dependency.                           │
└─────────────────────────────────────────────────────────────────┘
```

**Where a new language plugs in:**

1. **Language without GUI** (PHP, Ruby, Python in CLI mode, COBOL): nothing to do. Don't reference `dotnet::classes`. Don't add canvas host fns. The profile only needs `[builtins]` for the language's stdlib.

2. **Language with .NET-shaped classes** (any future C-derived .NET language, F#, …): reuse `dotnet::classes` directly — it's already in place. The language's profile sets `namespaces.use_dotnet = true`. The compiler's `register_dotnet_classes` runs at `compile()` start and installs the entire `Object → Form` chain as real callable globals.

3. **Language with its own GUI shape** (Dart Flutter `Canvas`, JS HTML5 `getContext('2d')`, Tkinter, …): write a new sibling submodule under `crates/vybe_compiler/src/emitter/<lang>/` containing wrapper classes whose method bodies use `MethodTarget::Body` sequences calling the SAME `vybe:gui::canvas*` host fns. Don't write a new canvas backend — Layer 1 + Layer 2 are framework-agnostic.

4. **Standalone Rust users of `vybe_widgets`**: they get all the widgets + `Canvas` trait + dialogs + everything in Layer 1 with zero VM dependency. The toolkit is shippable on its own.

The dotnet wrapper layer is the canonical example of pattern (2) and (3). When porting a fourth/fifth language, mirror its structure: a `classes/` submodule with a `mod.rs` (table + registration entry), `builder.rs` (chunk-building helpers), and one file per class family (object, control, form, buttons, …).

## Language-runtime adapters

A "runtime adapter" is a Rust function that emits bytecode composing existing WASM opcodes + spec-shape host fns to reproduce a language- or framework-specific runtime semantic. It is the right home for any helper whose behavior is tied to a specific language or framework rather than a spec primitive.

**Two adapter families** are wired today, with identical mechanics:

### PHP runtime adapter (`languages/php/src/emitter/`)

Every PHP runtime helper that doesn't 1:1 match an ECMA primitive — `date`, `strftime`, `strtotime`, `mktime`, `file_exists`, `filemtime`, `dir`, `str_pad`, `str_replace`, `ucwords`, `asort`, `array_chunk`, `base_convert`, `ctype_*`, etc. — lives here as `pub fn emit_<name>(chunks, current, argc, line)`. The PHP profile binds them via `emit = "common:php.<name>"`. **No JS polyfills, no PHP-specific host fns.** Adding a new helper means: add a Rust file (or a fn in an existing one) under `languages/php/src/emitter/`, wire it in `languages/php/src/emitter/dispatch.rs`, reference it in the profile. (The module is re-exported as `crate::emitter::php` for convenience.)

### .NET runtime adapter (`platforms/dotnet/src/emitter/core/`)

The .NET BCL surface that's shared across VB / C# / F# / any future .NET-shape language. Because it's a *platform* (multiple languages, prefix isn't a language name), it lives under `platforms/dotnet/src/emitter/` rather than `languages/<lang>/src/emitter/`. Same mechanics as PHP: `pub fn emit_<name>(chunks, current, line)` (or with `argc` for variadic), wired in `platforms/dotnet/src/emitter/dispatch.rs`, registered via the `"dotnet"` arm in `emitter::platform_emit_dispatch`, referenced via `common:dotnet.<name>` in profiles.

Currently active adapters:

| Adapter | Surface | Why it can't be a plain host call |
|---|---|---|
| `console_adapter` | `Console.Write` / `Console.WriteLine` / `Console.Print` | .NET prints booleans as `True` / `False` (capitalized) and `null` as `""`. JS `String(value)` differs. |
| `linq_adapter` | `First`, `Last`, `Skip`, `Take`, `ToList` / `ToArray` (identity), `Average`, `FirstOrDefault`, `Distinct`, `Aggregate(seed, fn)`, `OrderByDescending` | `Aggregate` swaps argument order vs. JS `reduce`; `FirstOrDefault` returns a typed default; `Average` divides Sum by length. None map to a single ECMA fn. |
| `array_adapter` | `Array.Clear` / `Array.Copy` / `Array.Resize` / `Array.Sort` / `Array.Reverse` / `Array.IndexOf` / `Array.Exists` / `Array.Find` / `Array.FindAll` / `Array.TrueForAll` / `Array.ConvertAll` / `Array.ForEach` / `List<T>.AddRange` | Range-mutate-in-place semantics; ECMA `Array.prototype.copyWithin` etc. don't match shape. The HOF variants (`Exists`, `Find`, …) compose existing `compiler_common::loops` emitters so the bytecode shape matches what instance-form `arr.some(...)` / `arr.filter(...)` produce. |
| `parse_adapter` | `int.Parse(s)` / `double.Parse(s)` / `bool.Parse(s)` | .NET `<T>.Parse` throws `FormatException` on invalid input per ECMA-335; JS `Number(s)` returns NaN silently. Adapter checks the result and throws a proper `Exception`-shape object so `e.Message` resolves correctly. |
| `datetime_adapter` | `DateTime.Now`, `DateTime.Parse`, `ToString(format)`, `AddDays`/`AddHours`/… | .NET format pictures (`"yyyy-MM-dd HH:mm:ss"`) differ from JS Intl. |
| `timespan_adapter` | `TimeSpan.FromSeconds`, `Add`, `TotalMilliseconds`, … | .NET-specific units / ratios. |
| `stringbuilder_adapter` | `StringBuilder.Append`, `AppendLine`, `Replace`, `Insert`, `ToString` | Mutable string buffer model; ECMA strings are immutable. |
| `string_format_adapter` | `String.Format("{0:N2}", x)` | .NET composite format strings. |
| `format_picture_adapter` | Numeric / date format pictures shared by `String.Format`, `ToString(fmt)`, etc. | Format-picture parser, used internally by other adapters. |
| `stream_io_adapter` | `StreamReader.ReadLine`, `StreamWriter.WriteLine`, … | .NET stream lifecycle differs from `wasi:filesystem` raw I/O. |
| `process_adapter` | `Process.Start`, `Process.GetCurrentProcess` | .NET process model. |
| `sockets_adapter` | `TcpClient`, `TcpListener`, … | .NET socket abstraction. |

**The rule:** when adding a new .NET BCL method that VB and C# both need, add a Rust adapter under `dotnet/core/` and reference it from BOTH profiles via `common:dotnet.<name>`. The walker should NOT bake .NET-specific lowering into the AST — that traps the implementation in one language. (See "Dispatch order — runtime collection intercept" below for the one current exception, where the dispatch pipeline forces a walker rewrite.)

### Other emitter adapter modules

`languages/dart/src/emitter/` (e.g. `is_empty`, `replace_first`, `list_first/last`, Dart `print`), `languages/fortran/src/emitter/` (`max`/`min` variadic, `len_trim`, `adjustl`), `languages/js/src/emitter/proxy_adapter`, `languages/cobol/src/emitter/`, `languages/ruby/src/emitter/` follow the same pattern. Each language owns its `emitter/dispatch.rs` exposing `pub fn dispatch(...)`, registered via the `LanguageDef::emit_dispatch` field in the plugin registry (`vybe_runtime::registry`). (All are re-exported as `crate::emitter::<lang>`.)

### Stdlib modules: adapters, NOT source preludes

When a language needs a stdlib module the runtime doesn't have (`os.path`, `random`, `math`, `string`, `datetime`, …), implement it as **adapters**, not as a pure-source "prelude" prepended to the program.

**The rejected anti-pattern (a hack — do not add more of these):** gating on `source.contains("modname")` and `parse_python_prelude(MODULE_PRELUDE)` to prepend ~150 lines of module source onto the `<script>` chunk. This parses + compiles the ENTIRE module on **every run** of any program whose source merely *contains* the substring (it fires on comments and variable names), so using one function drags in all of it — dead bytecode and real startup latency, every run.

**The pattern to use instead** — a pure/stateless module function is exactly a runtime adapter:

- Profile: `"os.path.join" = { emit = "common:python.ospath_join", min_args = 1, max_args = 255 }`.
- Dispatch (`languages/<lang>/src/emitter/dispatch.rs`): `"python.ospath_join" => os_path_adapter::emit_join(chunks, current, argc, line)`.
- Adapter (`languages/<lang>/src/emitter/os_path_adapter.rs`): `pub fn emit_join(...)` composes `ecma:string` / opcode ops with the args already on the stack. Emits **only at the call site of a function actually used** — no source parse, no substring gate, no bloat. String work follows `string_adapter.rs` (indexOf/lastIndexOf/substring/split); loop-heavy logic follows `math_adapter.rs` (`emit_block`/`emit_loop_s`/`emit_br_if`/`patch_loop`). Functions returning tuples use `common::tuples::emit_tuple`.
- For logic too gnarly to re-emit at every call site, use the **lazy shared-chunk** pattern (`repr_adapter.rs::ensure_py_repr_chunk`): create one named helper chunk on first use (`chunks.iter().position(|c| c.name == …)` → reuse, else build once) and `CALL` it. Still dispatch-gated and compiled once — an adapter, not a prelude.
- Module **constants** (`os.path.sep`) → string literals in the walker's `desugar_member_reads`; **`hasattr(mod, 'x')`** → a `py_module_surface` entry. Both are compile-time and clean — fine to keep.

**The one defensible use of a source prelude** is a genuinely **stateful class** module (Python `pathlib.Path` objects, `argparse.ArgumentParser`, `collections.Counter`/`deque`) — a class needs a class. Even then keep the prelude **thin** and have its heavy methods delegate to adapters/host, not inline everything. Prefer converting existing function-shaped preludes to adapters opportunistically; don't mass-rip working ones.

### libc platform adapter (`platforms/libc/`)

For languages that target the C runtime model (C, Go, Fortran, Rust/libc), all low-level emission is centralized under `platforms/libc/` rather than duplicated in each language walker. This is the platform equivalent of `platforms/dotnet/` — shared by any language where the runtime semantics match the libc ABI.

```
platforms/libc/
├── mod.rs             # pub mod declarations only
├── math_adapter.rs    # math.h — bare floor/ceil idents for profile dispatch
├── string_adapter.rs  # string.h read operations (strchr, strstr, etc.)
├── stdlib_adapter.rs  # malloc/calloc/free + atoi/atof
└── stdio_adapter.rs   # printf/fprintf → puts(sprintf(...))
```

Shared pointer/addressable-storage primitives live under
`crates/vybe_compiler/src/primitives/pointers.rs` and
`crates/vybe_compiler/src/primitives/addressable_storage.rs`.
Shared codepoint classification and complex-number primitives live under
`crates/vybe_compiler/src/primitives/codepoints.rs` and
`crates/vybe_compiler/src/primitives/complex.rs`.

**Pointer model — two kinds, both tagged at runtime:**

| Kind | Tag | Storage | Use for |
|---|---|---|---|
| Scalar cell | `{__ref_kind:"cell", __value:T}` | Existing `emitter/references.rs` | `&scalar_var`, `int x; int *p = &x;` |
| Array pointer (carray) | `{__ref_kind:"carray", __base:Array, __idx:i32}` | `primitives/pointers.rs` | `int *p = arr;`, `p++`, `p+n`, `p-q` |

**Walker usage pattern for `pointers.rs`:**

```rust
use vybe_compiler::primitives::pointers::{self, CARRAY_BASE_KEY, CARRAY_IDX_KEY};

// int *p = arr;  →  p = {__ref_kind:"carray", __base:arr, __idx:0}
pointers::make_carray_ptr(ident("arr"), lit_int(0))

// *p  →  p.__base[p.__idx]
pointers::carray_deref_read(p_expr)

// *p = val  →  p.__base[p.__idx] = val
pointers::carray_deref_write(p_expr, val_expr)

// p + n  →  {__ref_kind:"carray", __base:p.__base, __idx:p.__idx+n}
pointers::carray_advance(p_expr, n_expr)

// p - q  →  p.__idx - q.__idx
pointers::carray_diff(p_expr, q_expr)

// p++ (in-place mutation)  →  p.__idx = p.__idx + 1
pointers::carray_advance_inplace("p", lit_int(1))

// p-- (in-place mutation)  →  p.__idx = p.__idx - 1
pointers::carray_retreat_inplace("p", lit_int(1))
```

**Why libc, not `platforms/dotnet/src/emitter/`:** dotnet is a _language-platform_ (multiple languages sharing one runtime shape). libc is a _machine-platform_ (C, Go, Rust-without-stdlib, Fortran all share the same pointer/array/ctype semantics). Languages add their walker-level normalization on top; libc provides the emission primitives.

**What NOT to add here:** any ECMA-262 semantics, any PHP/Python/Ruby runtime helpers, any host fns. libc is pure AST constructor functions (returns `Expression` nodes) that downstream compilers emit as real WASM via the standard opcode/ecma/wasi path.

## Dispatch order — runtime collection registry & dotnet-adapter discriminator

The compiler's call-site dispatch (in `primitives/calls.rs`) orders resolution as:

1. **User class static method** (`MathUtils.Add(...)`) — runs first so user types aren't shadowed by built-ins
2. **Array methods (HOF dispatch)** — `[array_methods]` in profile (`Where → __array_filter`, `Select → __array_map`, …)
3. **Runtime collection registry vs. profile `[value_methods]`** — when the method name is registered as a Component-Model instance method on `dotnet.System.Collections.*` (e.g. `Add`, `Remove`, `Clear`, `Count`, `ContainsKey`, `Contains`), the dispatch normally defers to the runtime type registry in `vybe_platform_vybe::builtin_types` for type-aware behavior (List.Add → push, Dict.Add → set, HashSet.Add → set-add). **Discriminator:** if the profile's matched value-method overload routes to a shared `common:dotnet.<name>` adapter, the overload wins instead. This lets LINQ-style overloads (`Count(predicate)`, future `OrderBy(keySelector, comparer)`, …) bypass the per-type registry while leaving the existing type-aware `Add` / `Remove` / etc. behavior intact.
4. **Profile `[value_methods]` (general)** — including overloads by arity for non-`dotnet.*` emits
5. **Generic Member call** — fallback

The discriminator is a one-liner in `compile_call`:

```rust
let prefer_dotnet_adapter = match matched_value_method.as_ref().map(|d| &d.emit) {
    Some(BuiltinEmit::Common(name)) => name.starts_with("dotnet."),
    _ => false,
};
```

It only fires when (a) the profile has an overload that matches the call's arity AND (b) that overload's emit target is a `common:dotnet.*` adapter. The intent is: `dotnet/core/*_adapter.rs` is the canonical home for shared .NET BCL surfaces, so a profile entry pointing there is a deliberate "I want LINQ-style cross-language behavior, not per-type registry behavior" signal.

**Pattern to add a new LINQ / .NET surface that VB and C# both use:**

1. Write the bytecode in `platforms/dotnet/src/emitter/core/<name>_adapter.rs` (or extend an existing adapter, e.g. add to `linq_adapter.rs`).
2. Wire it in `platforms/dotnet/src/emitter/dispatch.rs` under `"dotnet.<name>"` (the platform's own dispatcher, registered via `emitter::platform_emit_dispatch`).
3. Add a `[value_methods]` overload to `csharp/profile` AND `vb/profile` referencing `common:dotnet.<name>`.
4. If the method name is in the runtime collection registry (Add, Count, etc.), the `prefer_dotnet_adapter` discriminator routes the arity-matched overload through your adapter instead. If the name isn't in the registry, the value-method dispatch picks it up directly.

No walker changes required for shared .NET behavior — keep walker rewrites for true language-syntax differences (e.g. PHP `M_PI` → `Math.PI`).

## Cross-language interop

All languages compile to the same bytecode and use the same struct layout. Interop is automatic:

- **Method aliases:** `emitter::canonical::cross_language_aliases` maps equivalent methods. Python `__str__` = JS `toString` = VB `ToString`.
- **Type stamping:** `set_type_id` + `__types` array. `instanceof` works across languages.
- **Exception names:** `emitter::errors::canonical_exception_name` maps language-specific names to canonical names.
- **Calling convention:** All functions are `ref_func` + `call` with arity. VB Sub and JS function produce identical bytecode.
- **Static inheritance:** Tracked via `PendingClass.statics` in the compiler. Child classes inherit parent statics at compile time.

---

## Debugging tools

When tests fail, use these tools. **The full reference is
[`documentation/debugging.md`](debugging.md)** — static dumps, execution tracing,
the interactive step debugger, VS Code/DAP, and hot reload. The essentials:

- **`cargo run -p vybex -- --dump-ast <file|project>`** — validate parser + walker shape before you touch the compiler. Large inputs print a top-level outline by default; set `VYBEX_DUMP_AST_FULL=1` for the full AST.
- **`cargo run -p vybex -- --dump <file|project>`** — inspect the generated chunks and bytecode summary without running the program. Add `--chunk <name|index>` to isolate a function or method.
- **`VYBE_TRACE=1 cargo run -p vybex -- --trace <file|project>`** — trace VM execution through the runtime path the CLI actually uses.
- **`cargo run -p vybex -- <file> --debug`** — interactive step debugger (breakpoints, stepping, `bt`, `locals`, `p <expr>`, watchpoints, hot `reload`). Works for **every** language since it's VM-level; the same REPL steps through your new language the day it compiles. See [`debugging.md`](debugging.md).
- **`cargo test -p vybe_language_<lang> --test <lang> --no-fail-fast`** — run the language suite you are changing.
- **`testrunner run tests/<lang>/<category>`** — run the *extracted* standalone files instead: no rebuild, and each case is a real source file you can hand to `-g` / `--dump-ast` / the real toolchain. See below.
- **`cargo run -p vybex -- --serve [:PORT] [ROOT]`** — debug request handling, project loading, and script execution in the web/server path. Use `--sandbox` to force restricted capabilities or `--no-sandbox` to keep full host access.
- **`vybe_runtime::debug::disassemble(chunk)`** — disassemble specific chunks from helpers or focused tests.
- **Rich VMError** — errors include call stack with chunk name + offset. Automatic for "not callable" and uncaught exceptions.
- **`__debug_dump(obj)`** — call from ANY language source to print all properties of an object to stderr. The compiler recognizes it as a built-in intrinsic.

### Where this language's tests live

Two locations, one of them generated:

| | path | what it is |
|---|---|---|
| **Source of truth** | `languages/<lang>/tests/<lang>/*.rs` | the Rust suites — write here |
| **Generated** | `tests/<lang>/<category>/<name>.<ext>` | one standalone source file per case |

A new language starts with the first only. Once its suites exist, extract them
so every case becomes a file you can debug directly — which matters most while
bringing a language up, when the failure is usually in the grammar or walker and
you want `--dump-ast` on the exact program that failed:

```sh
testrunner extract languages/<lang>/tests/<lang>/*.rs   # .rs suites are never modified
testrunner run tests/<lang>/<category>                  # no rebuild; prefer a category
vybex --dump-ast tests/<lang>/<category>/<name>.<ext>
```

`testrunner` needs one thing from the language: `crates/testrunner/src/emit/`
must know its file extension and, if its tests assert by comparing printed
output, a harness at `crates/testrunner/harness/<lang>/check.<ext>` written **in
that language** — the assertion helper is part of the test, so it is real source,
not Rust. A language whose tests assert by trapping (wast) needs no harness.

### Practical debug order when adding a language

1. Parse first: use `--dump-ast` to verify grammar precedence, statement boundaries, and walker normalizations before you look at the emitter.
2. Inspect emitted chunks next: use `--dump` and `--chunk` to confirm the primitives/profile/emitter path you expected is actually selected.
3. Trace runtime only after the AST and chunks look right: use `--trace` or `VYBE_TRACE=1` to find the first bad opcode or bad host call.
4. Use `--serve` when the failure only reproduces in request/response mode, import resolution, or bundled web roots.
5. Keep language syntax in `src/languages`, shared platform behavior in `src/platforms`, and reusable bytecode helpers in `src/emitter` while debugging. If the fix crosses those boundaries, stop and re-check the ownership.

### Expression eval for the debugger

The step debugger's *structural* features (breakpoints, stepping, `bt`, `locals`,
`p <name>` reads, `set`, watchpoints, hot reload) work for your language the day it
compiles — they operate on the VM, not on source syntax, so there is nothing to
add.

The one debugger feature that needs a per-language hook is **evaluating a typed-in
expression** (`p <expr>`, conditional breakpoints `b <loc> if <expr>`, and
`watch <expr>`). The evaluator (`crates/vybe_compiler/src/dynamic.rs`,
`debug_eval_expression`) parses the expression into the **common AST**, wraps it in
a common `FunctionDecl` whose parameters are the paused frame's locals, and
compiles+runs it in an isolated mini-VM — all language-agnostic. It only needs to
*parse an expression fragment* in your language:

- **Expression-oriented languages** (a bare expression is a valid program): nothing
  to do — `bundle_from_source(expr, language).prepared_module()` already yields a
  trailing `StmtKind::Expr` the evaluator lifts out. (JS, Python, PHP, Lua work this
  way.)
- **Statically-structured languages** (a bare expression is not a valid program):
  add a one-line arm to `eval_scaffold(language_name, expr)` in `dynamic.rs` that
  wraps the expression in your language's minimal parseable form, exposing it as
  `__vybe_frag`'s `return` — e.g. Go `func __vybe_frag() interface{} { return (…) }`,
  Java/C# `class …{ static object __vybe_frag(){ return (…); } }`. The shared
  `frag_return_expression` then lifts the expression whether `__vybe_frag` is a
  top-level function or a class/module method. (Go, C, Java, C# use this path.)

This lives entirely in the eval-service layer behind the debugger's eval hook —
it does not touch the compiler's normal path, and non-debug runs are unaffected.
See [`debugging.md`](debugging.md#language-support).

## Import resolution

The CLI (`crates/vybex/src/cli.rs`) resolves `import { x } from "./file.js"` by parsing the imported file and prepending its body to the main module. The walker produces `Module.imports` with `ImportKind::Named { path, names }`. The compiler ignores imports — resolution is the CLI's responsibility. Named exports (`export function`, `export let`) compile as regular declarations that become globals.

## Preflight checklist

Before marking a language as done:

- [ ] `grammar.pest` handles every statement type the existing parser handles
- [ ] `grammar.pest` handles every expression type including operator precedence
- [ ] `grammar.pest` keyword operators use atomic `@{}` rules for word boundaries
- [ ] `grammar.pest` expression chain is ~10-12 levels (not 18+)
- [ ] `walker.rs` maps every grammar rule to a common AST node — no panics/unimplemented
- [ ] `walker.rs` handles the `program` rule wrapper (unwrap `top.into_inner()`)
- [ ] `walker.rs` object literal methods prepend implicit `this` parameter
- [ ] `profile` has `[info]` section with name and extensions
- [ ] `profile` has ALL builtins from the existing compiler
- [ ] `profile` has value methods for instance method calls
- [ ] `profile` print function is mapped to `emit = "print"`
- [ ] `mod.rs` has `parse()` AND `profile_source()`
- [ ] Language registered in `languages/mod.rs` `all()` function (incl. `emit_dispatch` field)
- [ ] If the language has `common:<lang>.*` builtins: `languages/<lang>/src/emitter/dispatch.rs` exists and the central `emitter/dispatch.rs` was NOT touched
- [ ] Tests ported from old compiler to `languages/<lang>/tests/<lang>/`
- [ ] `cargo build -p vybex` compiles clean
- [ ] `cargo test -p vybe_language_<lang> --test <lang>` runs
- [ ] `cargo run -p vybex -- --dump-ast /tmp/test.<ext>` shows the expected AST shape
- [ ] `cargo run -p vybex -- --dump /tmp/test.<ext>` shows the expected chunks / emit path
- [ ] `cargo run -p vybex -- /tmp/test.<ext>` produces output

## What NOT to do

- Do NOT write a handcoded parser — use pest
- Do NOT create language-specific AST types — use the common AST
- Do NOT add language-specific code to the compiler OR to the central `emitter/dispatch.rs` — a language's emit adapters and its `dispatch.rs` live in `languages/<lang>/src/emitter/`, registered via the `LanguageDef::emit_dispatch` field. No `if name.starts_with("<lang>.")` in shared code; no `self.profile.name == "<lang>"` branches in `primitives/`.
- Do NOT hardcode builtins in Rust — put them in the profile TOML
- Do NOT create new crates for a language frontend — everything compiler-related goes in `languages/<lang>/src/`
- Do NOT simplify the walker — every construct the existing parser handles must be covered
- Do NOT guess builtins — read the existing compiler source for exact mappings
- Do NOT use direct host imports in the compiler — use `crate::emitter`
- Do NOT change emitter (`crate::emitter`) signatures — call existing functions on `&mut self.chunks[0]`
- Do NOT modify the VM (`vybe_runtime/`) — the VM contains ONLY WASM-compliant opcodes from real WASM proposals (MVP, GC, SIMD, Atomics, Memory64, String Builtins, Stack Switching). There were custom opcodes in the past — they have ALL been removed. ZERO custom opcodes remain. The VM is LAST RESORT — only modified to fix bugs in existing WASM opcode handling, never to add new behavior. Even WASM-compliant changes require explicit per-change user approval. Sort, LINQ, closures, everything goes through stdlib bytecode or `crate::emitter` using existing WASM opcodes.
- Do NOT add host functions to `platforms/{ecma,wasi,web,node,vybe}` without explicit per-function user approval. Host fns are ALWAYS ecma-shaped (`ecma:*` = ECMA-262/402/WebIDL) or wasi-shaped (`wasi:*` = WASI proposals). The ONLY non-spec namespace is `vybe:gui` for GUI controls/canvas/dialogs. NEVER add language-specific host functions — no `php:*`, no `python:*`, no `vb:*`, no `dart:*`. The default path is the emitter/adapter pattern — compose existing opcodes and host fns into bytecode.
- Do NOT modify unrelated crates when adding a language frontend — most work should stay in `crates/vybe_compiler/`, with `crates/vybex/` only changing if the executable/runtime surface needs it.
- Do NOT introduce dioxus-era concepts — no `SideEffect` enum, no global "side effect queue" the runner drains out-of-band, no `pending_*` `Vec<(String, String)>` field on `GuiState` for property changes / msgboxes / dialogs. The model is: VM calls a host fn → host fn talks to `vybe_widgets` directly → returns. Blocking dialogs block. Property writes go through `controlSetProperty` which mirrors into `GuiState::properties` AND fires the matching `WidgetEvent` from the widget's own `pending_events` (handled inside the widget's `handle_command` impl). Tests observe via host fn override, not by draining queues.
- Do NOT add framework-shaped types to `vybe_widgets` — `vybe_widgets::canvas` is HTML5-canvas-shaped, not .NET-shaped. No `Pen`, `Brush`, `Graphics` in `vybe_widgets`. Those are framework wrapper concerns that live in `emitter::dotnet::winforms::classes` (or a sibling emitter submodule for other frameworks).
- Do NOT take dependencies on `rfd` (or any GUI crate) directly from outside `vybe_widgets`. The toolkit owns dialogs. Other crates call `vybe_widgets::dialogs::*`.
- Do NOT special-case framework conventions in `compile_class` / the walker / the AST. Walker normalizations (implicit base call, implicit self, …) are general and reusable; they live in the language-specific walker, not the compiler.
- Do NOT put LINQ / .NET BCL methods in host runtime dispatch — they are .NET class methods and belong in `emitter::dotnet::core` adapters (referenced via `common:dotnet.<name>` in profiles), same as `Add`, `Remove`, `Sort`. See "Language-runtime adapters" above.
- Do NOT bake .NET-specific lowering into a single language's walker if the surface is shared with another .NET language — put the adapter in `emitter::dotnet::core` so VB and C# both pick it up by listing the same `common:dotnet.<name>` emit target. The dispatch in `primitives/calls.rs` (`prefer_dotnet_adapter`) routes profile overloads that target `common:dotnet.*` around the runtime collection registry, so even names already in the registry (`Count`, `Add`, …) reach the adapter when the arity matches an overload. Walker rewrites are for true syntax differences (PHP `M_PI` → `Math.PI`), not shared .NET BCL behavior.
- Do NOT add value_methods entries that conflict with type-registry methods (e.g. `keys`/`values` on Map). The value_methods dispatch runs BEFORE the type registry and intercepts calls meant for typed objects.
- Do NOT implement stdlib modules as source "preludes" (`source.contains("mod")` + `parse_python_prelude(MODULE_PRELUDE)` prepended to `<script>`). That parses+compiles the whole module on every run of any program whose source contains the substring — startup latency + dead bytecode. Pure/stateless module functions are ADAPTERS (`common:<lang>.<mod>_<fn>` → `dispatch.rs` → `emit_<fn>`), emitted only at the call site. Reserve thin preludes for genuinely stateful CLASS modules only. See "Stdlib modules: adapters, NOT source preludes" above.
- Do NOT use `host:vybe:runtime:*` for standard operations — if stdlib or `crate::emitter` provides it, use that. The only `vybe:*` namespace is `vybe:gui` for GUI controls/canvas/dialogs. All other host fns are `ecma:*` or `wasi:*`.
- Do NOT add language-specific host fns — every language is "X over JS". PHP, Python, VB, Dart, COBOL, Fortran, Ruby, Lua, Go all normalize to JS-shaped AST in their walker. The common compiler + ecma-shaped host fns handle everything. Language-specific runtime semantics go in emitter adapters (`languages/<lang>/src/emitter/`), not host functions.
