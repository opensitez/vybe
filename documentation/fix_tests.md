# Fixing Failing Tests — Vybe Field Guide

## How Vybe works (read this first)

Vybe compiles 15 languages to the same WASM bytecode VM. Every language is **"X over JS"** — its walker normalizes source code into a **common JS-shaped AST**, then a **single shared compiler** emits WASM bytecode. The compiler does not know which language produced the AST.

**Each language is its own crate.** The languages were extracted out of `vybe_compiler` into `vybe_language_<lang>` crates under `languages/<lang>/`, and each registers itself with the shared plugin registry (see "Plugin & registration model" below). Tests live with their crate at `languages/<lang>/tests/<lang>/` and run via `cargo test -p vybe_language_<lang> --test <lang>`.

```
Source code (JS, PHP, Python, VB, C#, Dart, Ruby, COBOL, Fortran, Lua, Go, C, Pascal)
    │
    ▼
┌──────────────────────────────────────────────────────────────────┐
│  GRAMMAR (grammar.pest)                                          │
│  Fully compliant PEG grammar. Parses every valid construct.      │
│  Output: pest parse tree.                                        │
└──────────────────────┬───────────────────────────────────────────┘
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  WALKER (walker.rs) + NORMALIZER (normalize.rs if present)       │
│  Normalizes ALL language-specific idioms into JS-shaped AST.     │
│  PHP count($arr) → arr.length, VB Len(s) → s.length, etc.       │
│  After this step, the AST is language-agnostic.                  │
│  Output: vybe_ast::Module                              │
└──────────────────────┬───────────────────────────────────────────┘
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  PROFILE (profile TOML)                                          │
│  Maps builtin names → emit strategies. Tells the compiler how    │
│  to emit bytecode for this language's builtins and methods.       │
│  Missing profile entries = "undefined" / "not callable" errors.  │
└──────────────────────┬───────────────────────────────────────────┘
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  COMPILER + EMITTER (shared — one for ALL languages, ONE module) │
│  crates/vybe_compiler/src/primitives/*.rs                          │
│  Emits standard WASM opcodes only. Uses ecma:* host fns.         │
│  Output: Vec<Chunk> (WASM bytecode)                              │
└──────────────────────┬───────────────────────────────────────────┘
                       ▼
┌──────────────────────────────────────────────────────────────────┐
│  VM (vybe_runtime) + HOST PLATFORMS                            │
│  (platforms/{ecma,wasi,web,node,vybe}, each a Plugin)           │
│  WASM-compliant execution. ZERO custom opcodes.                  │
│  Host fns: ecma:* / wasi:* / web:* / node:* (exception: vybe:gui)│
│  OFF-LIMITS without explicit per-change user authorization.      │
└──────────────────────────────────────────────────────────────────┘
```

---

## HARD RULES — read before touching ANY code

These are non-negotiable. Breaking them has caused reverts, regressions, and lost days.

1. **NEVER add custom opcodes** to `vybe_runtime/`. The VM contains ONLY real WASM proposal opcodes. Zero custom opcodes remain — all past ones have been removed.
2. **NEVER modify the VM** without explicit per-change user authorization. The VM is last resort — only touched to fix bugs in existing WASM opcode handling.
3. **NEVER add host functions** to the host-platform crates (`platforms/{ecma,wasi,web,node,vybe}/`) without explicit per-function user authorization. Host fns are always ecma-shaped (`ecma:*`), wasi-shaped (`wasi:*`), `web:*`, or `node:*`. The only non-spec namespace is `vybe:gui`. No language-specific host namespaces (`php:*`, `python:*`, `vb:*`).
4. **NEVER add polyfills** — no `__vybe_*` helpers in `.js` files, no `polyfills/php_*.js`. Runtime semantics go through emitter adapters.
5. **NEVER implement a stdlib module as a source "prelude"** — no `source.contains("mod")` gate + `parse_python_prelude(MODULE_PRELUDE)` that prepends the whole module to the `<script>` chunk. That parses+compiles ~all of it on every run of any program whose source merely contains the substring (fires on comments/var names) — startup latency + dead bytecode. Pure/stateless module functions are ADAPTERS (`common:<lang>.<mod>_<fn>` → `dispatch.rs` → `emit_<fn>`), emitted only at the call site (template: `math_adapter.rs`, `string_adapter.rs`; lazy shared chunk: `repr_adapter.rs::ensure_py_repr_chunk`). Thin preludes are acceptable ONLY for genuinely stateful CLASS modules (a class needs a class).
6. **NEVER fix a test to mask a real bug.** If a test is correct per spec and Vybe fails it, fix Vybe (with auth), not the test.
7. **NEVER emit raw opcodes** in `primitives/mod.rs` or a walker. Use the shared emit helpers (`common::ops`, `common::collections`, … — see below).
8. **NEVER run git commands** unless the user explicitly asks.
9. **NEVER modify working shared code** (the host-platform crates, `vybe_runtime`, compiler core) to fix one language's tests — it will regress other languages and you cannot verify without running all suites.
10. **NEVER run cross-language test suites** unless the user explicitly asks. Only run the specific test or file you're working on.
11. **Follow the user's instructions exactly.** Do not expand scope. Do not start investigating/fixing things that weren't asked for.

---

## File map — where to look

```
languages/<lang>/                        ← one crate per language: vybe_language_<lang>
├── src/
│   ├── grammar.pest        ← PEG grammar (parse errors, missing syntax)
│   ├── walker.rs           ← pest tree → common AST (normalization bugs)
│   ├── normalize.rs        ← optional normalizer (Lua, etc.)
│   ├── normalize_class.rs  ← per-language class → NormalClass IR
│   ├── profile             ← TOML: builtins, value_methods, intrinsics, compiler flags
│   ├── tree_register.rs    ← optional namespace-tree registration
│   ├── lib.rs              ← parse() + profile_source() + register() + `struct Plugin`
│   └── emitter/
│       ├── dispatch.rs     ← routes common:<lang>.* to adapter fns
│       └── *_adapter.rs    ← language-specific runtime → bytecode
└── tests/<lang>/           ← THIS language's test suite (helpers.rs + test_*.rs + main.rs)

platforms/<platform>/                    ← one crate per target platform
│   EMIT platforms (compile-time codegen):
├── libc/src/emitter/       ← C runtime model: stdio, string.h, ctype, math
├── dotnet/src/             ← .NET surface (emitter/ + winforms/) shared by VB & C#
├── plib/src/               ← Pascal component library (GUI, etc.)
├── flutter/src/            ← Flutter widgets adapter
├── wasm/src/               ← WASM codec / disassembler
│   HOST platforms — each exposes ONE `Plugin` (same type as languages);
│   its `init` registers that platform's host fns, capability-gated:
├── ecma/src/               ← ecma:* runtime + ecma_globals (constructor↔proto + Intl wiring)
├── wasi/src/               ← wasi:* (clock/console/env/random/fs/filesystem/io/http/sockets/sql/crypto)
├── web/src/                ← web:* (crypto/URL/fetch/dom-parser)
├── node/src/               ← node:* (fs/os/path/process/http/child_process)
└── vybe/src/               ← vybe:gui (controls/canvas/drawing/gui_state); one Plugin owns its GuiState (feature `gui`)

crates/vybe_compiler/src/                ← the ONE shared compiler, the SHARED EMITTER, and the eval layer
├── primitives/              ← ONE module. It IS the emitter — there is no `emitter/` dir and no emitter crate.
│   ├── mod.rs             ← conductor
│   ├── calls.rs classes.rs statements.rs operators.rs link.rs …   ← `impl Compiler` AST walkers
│   ├── ops.rs collections.rs dict.rs strings.rs math.rs loops.rs  ← SHARED cross-language bytecode
│   │   errors.rs functions.rs generators.rs io.rs convert.rs         helpers: free fns over `&mut Chunk`,
│   │   tuples.rs sprintf.rs random.rs delegates.rs …                 AST-free, standard WASM only
│   ├── expressions.rs events.rs reflection.rs  ← topics with BOTH halves in one file
│   ├── instructions/      ← low-level opcode recipes (core_wasm.rs, host.rs)
│   ├── dispatch.rs        ← routes common:* (shared cats direct; lang/platform prefixes via registry)
│   ├── bundle.rs runtime_helpers.rs polyfills/  ← bundled stdlib chunks + polyglot polyfills
│   └── platforms.rs       ← drives the ONE plugin registration loop for a VM (names no platform)
├── languages/mod.rs       ← thin facade over vybe_runtime::registry (all()/find_by_name/...)
├── dynamic.rs             ← DYNAMIC COMPILE / eval / PHP include / JS Function
├── host_imports.rs        ← ESM host-import installer
└── build.rs               ← generates `extern crate` per plugin dep so link-time registration works
    NOTE: vybe_compiler depends on NO emit-platform crate. Platforms depend on IT.

crates/vybe_runtime/       ← VM = THE RUNTIME (+ capabilities, js_builtins) AND the plugin SDK:
│                             framework.rs (Plugin trait, Framework, PluginEntry, the ONE init loop),
│                             registry.rs (LanguageDef/PlatformDef/hooks), namespaces.rs (the tree),
│                             profile.rs, class_normalize/ — OFF-LIMITS w/o per-change auth
crates/vybex/               ← thin launcher/CLI/server ONLY — OFF-LIMITS unless explicitly told
```

**`vybex` is off-limits unless the task explicitly names it.** It is a leaf: it wires up
the CLI, server, and GUI launch and calls into `vybe_compiler` code. Nothing
should depend on it, and language work never edits it. The dynamic-compile / eval service
used to live here — it now lives in `vybe_compiler` (`dynamic.rs`) so language crates and
their tests can reach it without depending back up on the shell.

---

## Plugin & registration model

There is exactly **ONE plugin type**: `vybe_runtime::Plugin`. Every capability provider — a
source **language**, a host **platform**, an emit platform — is that one type. A provider has
**one** `struct Plugin` (per crate) whose `init(&mut Framework)` registers **whatever
capabilities it has** — a language, host functions, gui, an emit dispatcher, a namespace tree,
any mix. There are **no** per-role plugin types (no `HostPlugin`/`GuiPlugin`/`EcmaPlugin`), and
what a plugin registers are **capability descriptors** (`LanguageDef`, `PlatformDef`, tree
`Type` nodes), NOT plugins. The compiler never names a concrete language or platform crate.
*(Why one type: so all plugins share one registry, are driven by one generic loop, and — the
end goal — can be `dlopen`'d uniformly as `.so`/`.dll`. `Plugin::init` is the in-process form
of the future `extern "C"` entry point; `Framework` is the versioned registrar / ABI surface —
grow it with a new `register_*` method per capability, never a per-role type.)*

**There is no plugin list in code, anywhere.** Each plugin crate ends with
`vybe_runtime::register_plugin!(Plugin);`, which submits it to the registry at LINK time; a
`build.rs` reads its own Cargo.toml and emits one `extern crate` per plugin dependency, because
Rust drops an rlib nothing references. So a binary's plugin set is decided by its **Cargo
dependencies** — adding or removing a plugin is a Cargo edit and touches no source file. One
loop, `vybe_runtime::init_all_registered(vm, caps)`, drives whatever is linked.

- **Host platforms** (`platforms/{ecma,wasi,web,node,vybe}`) each expose **one** `Plugin`;
  its `init` registers that platform's host fns on the VM. `required_capability()` (or
  internal `fw.granted(cap)` for multi-cap platforms like wasi/node) gates them. `finalize`
  is the second phase, for work that must observe a fully-populated VM — ecma wires
  `globalThis` + constructor↔prototype + `Intl` there, since it resolves host fns other
  plugins registered by index. There is no separate host crate: `vybe_host` was split into
  these platform crates, and the old hand-written `register_with_capabilities` is gone.
  Stateful plugins (vybe/gui) keep their state **beside** the plugin, not inside the value —
  the registry holds `&'static dyn Plugin`, so `Plugin::with_gui()` installs a fresh
  `GuiState` and `gui_state()` reads the shared handle back.
- A language crate's `src/lib.rs` exposes `pub fn register()` and `pub struct Plugin`
  (`impl vybe_runtime::Plugin`). `register()` calls
  `vybe_runtime::registry::register_language(LanguageDef { name, parse, profile_source,
  emit_dispatch, normalize_class, register_tree })` — a **capability descriptor**, not a
  plugin — and, if the language needs core hooks, `register_hooks(name, LanguageHooks { … })`.
  A language plugin could *also* register host fns / an emitter in the same `init`.
- `LanguageHooks` carries the small set of behaviors the shared compiler must call *back*
  into a language for: `value_eq` (Python), `relational_compare` / `str_getcsv` /
  `normalize_source` (PHP), `proxy_*` (JS), `parse_eval` (JS eval-string parsing), etc.
  Everything optional — a language registers only what it has.
- The shared compiler resolves everything by name through the registry
  (`crate::languages::find_by_name`, `emit_dispatch_for`, `registry::hooks(name).<hook>`),
  so a failing dispatch usually means the language's `register()` wasn't called or a hook
  field is `None`. Test helpers call `register()` once (guarded by `Once`) before compiling.

The same inversion applies to platforms: a platform registers a `PlatformDef` (emit dispatch,
tree registration, component descriptors) and its namespace-tree `Type` nodes, and the compiler
resolves both by name/prefix through the registry. A GUI platform's whole compiler-side
integration is one profile line — `type_scopes = ["flutter"]` / `["plib"]` / `["dotnet"]` —
which both scopes member resolution and mounts the root ambiently, so unqualified platform
names construct through the one common-resolver `Ctor` path. No compiler-side registration
pass, no per-platform hook.

## Dynamic compilation & eval — lives in `vybe_compiler`

`eval`, `new Function`, and PHP `include` compile a **source string at runtime** and run it.
That code (`vybe_compiler/src/dynamic.rs` + `host_imports.rs`) sits in `vybe_compiler`
because it depends on the compiler — it *cannot* live in `vybe_runtime` (that would cycle)
and must **not** live in `vybex` (that made language tests depend up on the shell). The
`RuntimeCompilerService` there is what the php/python/js test helpers use to run programs
that may `include`/`eval` mid-execution. It reaches JS-specific eval parsing through the
`parse_eval` hook rather than naming the JS crate.

---

## Step 0 — Reproduce in isolation

This is the single highest-leverage habit. A failing suite test is a bad debugger; a
2-line standalone program is a good one.

**Check `testrunner` first — the standalone file may already exist, and then
there is nothing to copy into `/tmp` at all.** The Rust suites have been
extracted into per-test source files under
`tests/<language>/<category>/<name>.<ext>`, and the failure slug *is* the path:

```bash
testrunner run tests/go/json_marshal 2>/dev/null | grep FAILED
# test go/json_marshal/marshal_bool_true ... FAILED
#      └─ tests/go/json_marshal/marshal_bool_true.go

vybex -g         tests/go/json_marshal/marshal_bool_true.go   # step through it
vybex --dump-ast tests/go/json_marshal/marshal_bool_true.go
go run           tests/go/json_marshal/marshal_bool_true.go   # ground truth
```

That is steps 1 and 3 below already done, on the *exact* program the suite ran —
no transcription risk, and the step debugger works on it directly. Carve a
smaller repro only when the extracted file is too big to reason about.

**Name a category, not a suite.** A target may be `tests/<lang>` (whole suite),
`tests/<lang>/<category>` (one category) or a single file, and you may pass more
than one — including across languages. A suite is thousands of tests and minutes
of wall clock; a category is seconds:

```bash
testrunner run tests/go/json_marshal tests/go/bufio_io    # two categories
```

Output is `cargo test` shaped (`--progress` for a live table), coloured on a
terminal and plain when piped, so a redirected run stays diffable. It adds one
verdict cargo lacks — `TIMEOUT`, because a hang never produced an answer to be
wrong — plus a one-shot `warning: <slug> still running after 10s` that is not a
verdict at all. It is slower than `cargo test` on a full suite but it does not
hang: the deadline is 60 s and the rest of the run continues.

See [`testrunner.md`](testrunner.md) and
[`debugging.md`](debugging.md#getting-a-debuggable-file-out-of-a-failing-test).

1. **Carve out a minimal standalone file** that reproduces the wrong output:
   ```bash
   cat > /tmp/r.js <<'EOF'
   // the smallest thing that misbehaves
   console.log("hello");
   EOF
   cargo run -q -p vybex -- /tmp/r.js 2>/dev/null
   ```
   For C, also compare against the real compiler:
   ```bash
   cargo run -q -p vybex -- /tmp/r.c 2>/dev/null | od -c    # Vybe output (exact bytes)
   /usr/bin/cc /tmp/r.c -o /tmp/r && /tmp/r | od -c          # ground truth
   ```
   `od -c` is essential when the diff involves whitespace, NULs, or embedded newlines.

2. **Bisect by toggling ONE variable at a time.** Build a ladder of variants that differ
   by exactly one axis and find where it flips:
   - loop vs. unrolled (same calls, no loop)
   - through a function call vs. inline
   - with closure vs. without
   - object property vs. plain variable
   - with inheritance vs. flat class

3. **Compare against the real runtime** (`node`, `php`, `python3`, `/usr/bin/cc`) to settle
   "is the test even right?" Always pipe through `od -c` and diff — never eyeball.

4. Only after the repro is minimal and the layer is known do you open a source file.

---

## Harness field notes — read before blaming the language

Extracted tests are one file each at `tests/<lang>/<category>/<name>.<ext>`,
and **the verdict is the exit code**. Every rule below was paid for once.

### The differential is the first check, not the last

**If the real runtime passes far fewer tests than Vybe does, the HARNESS is
wrong — not the language.** Measured twice on Dart: 2/31 then 9/31 under the
real SDK against Vybe's 23/31, and both times it was my emitter. After fixing
both it was 30/31 in *both*, with the single failure DIFFERING between them —
which is the signal you actually want:

- fails only on Vybe → a real Vybe bug
- fails only on the real runtime → the expectation is wrong, fix the test

Reference toolchains present: `node`, `python3`, `php`, `go`, `dart`,
`/usr/bin/java`, `cobc`, `gfortran`, `fpc`, `wasmtime`.

```sh
testrunner run tests/dart/<category> --runtime "dart run"
testrunner run tests/cobol/<category> --runtime cobc      # see limits below
```

### Prove the harness cannot FALSE-PASS

A harness that reports success when the assertion never ran is worse than no
harness. Corrupt a passing test's expected value and confirm it fails **in both
runtimes**:

```sh
sed 's/want-value/WRONG/' t.rb > bad.rb
vybex bad.rb; echo $?      # must be non-zero
ruby  bad.rb; echo $?      # must be non-zero
```

Two real false passes this caught:
- Ruby's `return if got == want` guard returned unconditionally, so a failing
  test exited 0.
- COBOL's first harness used `STOP RUN WITH ERROR STATUS`, which is **not in
  the grammar** — every check was a parse error, so tests "failed" for a reason
  unrelated to their assertions.

### `--cold` is the warm-pool invariant

Warm and cold must agree on every verdict. When they disagree, believe it. The
warm worker once read only the `Result` and ignored `vm.pending_exit_code`, so
a Python `sys.exit(3)` **passed** under the default pool and failed under
`--cold`, under `vybex <file>` and under `python3`. Re-run `--cold` after any
change to the runtime, not just to the harness.

### An early `exit` can skip a trailing check

Harnesses that compare the whole output at the end (php, ruby, java, dart,
pascal) put one check after the program. If the program calls
`exit`/`die`/`Halt`, that check never runs and the test asserts **nothing**
while still looking green. Only 8 tests corpus-wide do this (7 php, 1 pascal);
the per-print languages are immune because their checks are inline.

For PHP the fix is `register_shutdown_function`, which real php honours.

### Language-specific things that cost a debugging session

- **Fortran** — compare VALUES, never printed text. `print *` pads
  (`"           8"`) and writes logicals as `T`/`F`, so a textual comparison
  fails under gfortran on formatting alone. Emit `(x) /= 8`,
  `abs((x) - 3.5) > 1.0e-6`, `(x) .neqv. .true.`
- **Dart** — `await` on a `void` expression is a compile error in real Dart
  even though Vybe accepts it. And `compile_ok` cases have no `main`, which
  `dart run` refuses outright (exit 253) while `vybex -d` does not need one.
- **Pascal** — `WriteLn` takes any type and Vybe has no `WriteStr`, so use an
  **overload set** (`__vs(v: string)`, `__vs(v: integer)`, …) and let the
  compiler pick. Add `{$mode delphi}` + `uses SysUtils` so one file runs on
  both fpc and vybex.
- **COBOL** — free-format `*>` comments (fixed-format `      *` in column 7 is
  rejected by both). `cobc` rejects a base filename over **31 characters**
  (measured: 31 passes, 32 does not; the directory is not counted). Repeat
  `DELIMITED SIZE` after EVERY operand — Vybe does not propagate a single
  trailing delimiter back over the preceding ones the way cobc does.
- **Java** — the harness must be a static member INSIDE the test's class, since
  Java has no top-level functions.
- **Ruby** — `print` inside any conditional executes unconditionally, so the
  harness carries no diagnostic at all until that is fixed.

---

## Step 1 — Is the test wrong?

Before touching any code, verify the expected value is correct per the relevant spec.

**Common wrong-test patterns found in the field:**

| Language | Symptom | Correct answer |
|---|---|---|
| JS | `Math.round(-3.5)` | `-3` (rounds toward +∞ per ECMA §21.3.2.28) |
| JS | `[, secondItem]` destructuring on 1-element array | `undefined` — skips index 0 |
| JS | `Number.isInteger(-0)` | `true`; `Number.isInteger(Infinity)` → `false` |
| JS | `arr.flat(2).join(",")` — nested arrays stringify | `[4,[5]].toString()` → `"4,5"` |
| VB | `Integer / Integer` | always truncates, not rounds |
| PHP | `intdiv(-7, 2)` | `-3` (truncates toward zero, not floor) |
| Any | `if (cond) break; push(n)` — break fires BEFORE push | one fewer element |

Check the spec (ECMA-262, PHP manual, VB.NET docs) or the real runtime before changing code.

---

## C# field notes — most failures are frontend gaps

Most current C# failures are **unimplemented C# language behavior in the grammar,
walker, or normalization layer**, not VM bugs and not host gaps. Treat C# as
**C# over JS over WASM**:

```
C# syntax
    → grammar.pest parses it
    → walker.rs / normalize.rs lowers it to the common JS-shaped AST
    → shared primitives/emitter produces WASM
```

So for C# tests, start here:

1. `languages/csharp/src/grammar.pest`
   - Use when the test fails to parse or a construct is absent from the tree.
   - Typical cases: missing pattern syntax, query expression forms, modifiers,
     generic constraints, interpolated string variants, or newer C# operators.

2. `languages/csharp/src/walker.rs`
   - Use when parsing succeeds but the AST is wrong, incomplete, or too
     C#-shaped for the shared compiler.
   - Typical cases: `using`/namespace resolution, `System.*` member chains,
     `this`/`base`, properties/indexers/events, pattern matching, collection
     initializers, LINQ, lambdas, null-coalescing/null-conditional operators,
     and C# string/array helpers that should become JS-shaped member calls.

3. `languages/csharp/src/normalize.rs` if present, then
   the C# profile
   - Use when the AST shape is mostly right but needs a final canonical form or
     a builtin/value method mapping.
   - Missing profile entries usually show up as `undefined is not callable`;
     wrong walker lowering often shows up as `undefined`, `[object]`, or a
     method/property access on the wrong receiver.

Only move below the C# frontend after a minimal repro proves the common AST is
correct and the failure is genuinely in shared primitives/emitter behavior. Do
not fix C# gaps by adding host functions, VM behavior, JavaScript polyfills, or
language-specific checks in shared compiler code.

---

## Step 2 — Categorize the failure signal

```
left: [...]    ← what Vybe actually produced
right: [...]   ← what the test expects
```

| Signal | Likely cause | Where to look |
|---|---|---|
| `left: []` (empty output) | Crash, compile error, unresolved import, or async not awaited | Walker / grammar / profile |
| `panicked ... compile failed` | Parse error or walker bug | Grammar.pest or walker.rs |
| `panicked ... RuntimeError: undefined is not callable` | Function not found — missing profile entry, wrong variable name, or closure not captured | Profile builtins, walker normalization, closure emit |
| `panicked ... RuntimeError: null is not callable` | Function resolved to null — wrong dispatch, missing method binding | Compiler calls.rs, class emit |
| `left: ["[object]"]` | Method not found on object, wrong property key | Walker normalization |
| `left: ["[function <lambda>]"]` | Returned the function instead of calling it | Walker / compiler |
| `left: ["undefined"]` | Property/variable not found, wrong key emitted | Walker normalization, profile |
| `left: [wrong number]` | Wrong arithmetic, wrong coercion, wrong operator | Compiler expressions.rs or emitter |
| `left: [correct, correct, wrong]` | Off-by-one, ordering, or precision issue | Walker or compiler |
| `left: ["NaN"]` | Type coercion failure, missing conversion | Walker normalization or emitter convert |
| `[vybe] host call arity: …` on stderr | A **declared** host fn is called with the wrong argument count — see below | The emitter that built the call, or the declaration |

### `[vybe] host call arity:` — read it, never silence it

```
[vybe] host call arity: web:html:setValue declares 3 parameter(s), called with 2 (line 88)
```

Host functions registered through `HostFnDecl` carry a Component Model signature,
and `Chunk::emit_call` checks every emitted call against it at **compile** time.
This line means one of exactly two things, and they need opposite fixes:

1. **The call site is wrong.** An emitter forgot an operand. The missing argument
   arrives as `Value::Null`, so the symptom shows up far away — an empty string
   where text was expected, a control that never got its parent. Fix the emitter.
2. **The declaration is wrong.** The signature was written from the IDL rather than
   from what the closure actually reads. Fix the declaration in
   `platforms/<p>/src/*.rs`.

**Never trim a declaration to match a wrong caller** — that converts a loud
compile-time finding into the silent runtime bug it was built to catch.

Two things this check does NOT tell you:

- **Silence is not coverage.** The check fires at emit time, so a clean run only
  proves the paths that program actually compiled. A function with no caller at all
  never warns. If you need to know a route exists, read the emitter or disassemble —
  do not read absence of warnings as proof.
- **Undeclared functions are never reported.** `declared_host_arity` answers `None`
  for them, which means UNKNOWN, not zero. Coverage as of this writing: **172 of
  ~1780** host functions — `platforms/web/src/html.rs` (38, the reference shape
  for a DOM-style module), `window.rs` (13), `animation.rs`/`crypto.rs` (3 each),
  3 of `timers.rs`, `platforms/wasi/src/io.rs` (23, the reference shape for a WIT
  resource module), `fs.rs` (26 of 31), `clock.rs` (15), `random.rs` (8 of 9),
  `env.rs` (4) and `sql.rs` (36 of 42). All of `ecma:*`, `node:*` and `vybe:*`,
  and the rest of `wasi:*` — `sockets.rs` (123), `http.rs` (71),
  `filesystem.rs` (38), `crypto.rs` (23), `console.rs` (10) — are undeclared.

**Pick the next module by REACHABILITY, not by size.** The check fires at emit,
so declaring a function nothing calls buys documentation and no verification.
Measured call sites per interface, against registrations: `wasi:sql` 31/42,
`wasi:sockets` 16/123, `wasi:http` 5/71, `wasi:filesystem/types` 0/38,
`wasi:crypto` 0/23. `sql.rs` was worth four times what `sockets.rs` is worth
despite being a third the size.

**A silent run proves inertness; a PREDICTED warning proves the mechanism.**
Before trusting a batch of new declarations, find one call site you expect to be
wrong, and confirm the warning fires there *and nowhere else*. The `wasi:cli/environment`
batch had one: `get-environment` is `func() -> list<tuple<string, string>>` — no
argument, in both the 0.2 and 0.3 WIT — but VB's `Environ$(name)` lowers to it
with the key pushed (`primitives/builtins.rs`, the `"environ"` arm). Declaring
the truthful `0` made `tests/vb/vb_interaction_environ_command` report
`declares 0 parameter(s), called with 1`, and no other file in the sweep said
anything. That is the check working, and it is also a real bug: the host ignores
the key and answers the whole list, so VB `Environ` cannot return one variable.

**For a VYBE-DEFINED surface the authority is the descriptor, not a WIT.** Only
five of `sql.rs`'s 42 functions are in `proposals/wasi-sql/wit/`; the rest of
`wasi:sql/types` is the ADO.NET shape that `documentation/dotnetdb.md` and
`complianceplan.md` Phase 9 (`vybe:data → wasi:sql`) deliberately put there, and
flat `wasi:sql` is the pre-resource shape PHP and Python still use. Extending a
`wasi:*` interface this way is the design, not drift — but do not write
WIT-citing comments for names that are not in a WIT; a fabricated citation is
worse than none. What routes those calls is the .NET component descriptor
(`platforms/dotnet/src/emitter/core/component_classes_data_drawing.rs`), and
each of its three forms converts to a host argc differently:

- `MethodDef::new(name, N, MethodBody::HostCall(..))` → **argc N+1**. The lookup
  in `platforms/dotnet/src/emitter/mod.rs` matches `method.arity == arg_count`
  and `primitives/calls.rs` pushes the receiver first (`arg_exprs.len() + 1`).
- `ConstructorDef::new(N).with_backing(..)` → **variable, do not declare**.
  `lookup_type_ctor_target` matches the class name with NO arity filter and
  `primitives/expressions.rs` then emits `args.len()`, so the guest picks.
- Property reads reach `primitives/expressions.rs` which emits
  `emit_host_call(idx, 1)` unconditionally — so anything also exposed as a
  property is argc 1 no matter what its descriptor says. Check for a
  `PropertyDef` before declaring anything at argc > 1.

Types for such a surface should be `ValType::Any` where the value is a plain
property-bag `Object`. Writing `Own`/`Borrow` would claim resource handles that
do not exist.

**`wasi:filesystem` and `wasi:filesystem/types` are two different surfaces.**
`fs.rs` registers the bare `wasi:filesystem` with camelCase names (`openFile`,
`readFile`) — that is what emitters actually import, with 15 measured call sites.
`filesystem.rs` registers the WIT-shaped `wasi:filesystem/types`
(`[method]descriptor.open-at`) which has **no** compiler call sites at all.
Declare the one that is called. Related: `filesystem.rs` also re-registers 7
`wasi:io/streams` functions that `io.rs` registers again afterwards — and
`lib.rs:35` orders them that way on purpose, so those 7 are shadowed and dead.

**`wasi:*` is where a declaration says the most.** Its registrations are already
named in WIT's resource syntax — `"[method]input-stream.read"`,
`"[resource-drop]pollable"` — which is resource membership encoded in a STRING.
`resource_member` turns that convention into something checkable, and `own` vs
`borrow` then states what the prefix never could: a destructor CONSUMES its
handle while every other method borrows it. 233 registrations across
`io`/`filesystem`/`http`/`sockets`/`sql` carry those prefixes.

When declaring WASI, **take the signature from the WIT, not from the closure.**
Several closures ignore their `self` handle — `flush`, `pollable.ready` and every
`[resource-drop]` read no arguments at all, because this host has one instance
and does not need the handle to find it. The call sites still pass it
(`pollable.block` is emitted with argc 1, `blocking-read` with 2), so the closure
is a LOWER BOUND on the contract, not the contract.

**Measuring call-site arity: four forms, and three ways to get it wrong.** The
argc a declaration must match is spread across `call_import(chunks, current,
"mod", "name", ARGC, line)`, the free-function `emit_host_call(chunks, current,
"mod", "name", ARGC, line)`, a `let idx = …add_import("mod", "name")` followed
by `emit_call(idx, ARGC, line)`, and the declarative
`HostTarget::new("mod", "name")`. A scan that models only the first and third
under-reports badly. Three traps, each of which produced a wrong number before
it was caught:

- **Scan `platforms/` too.** `crates/vybe_compiler` and `languages` are not the
  whole caller set — `platforms/dotnet` and `platforms/libc` hold the emitters
  for `wasi:sockets`, `wasi:filesystem` and much of `wasi:clocks`. Omitting them
  makes a heavily-called module look uncalled.
- **`import(` spans lines.** When a binding's `self.import(\n "mod",\n "name")`
  is invisible to a single-line pattern, a later `emit_call(idx, N)` gets
  attributed to whatever that variable was bound to *earlier in the file* — which
  is how `wasi:cli/stdout.write-via-stream` was once credited with an argc that
  belonged to `output-stream.blocking-write-and-flush`. Resolve each call to the
  NEAREST PRECEDING binding, and treat single-site attributions in match-heavy
  files as weak.
- **`emit_call(idx, 0u8, line)` is argc 0.** A pattern that only accepts bare
  digits records it as a variable and invents a phantom "arity varies here".

A genuinely variable argc — `emit_vb_get(chunks, current, argc, line)` forwarding
the guest's own argument count — is real signal: it means the arity is
guest-driven, so the check reports a *guest* mismatch rather than an emitter bug.
Verify those empirically by running the suite with stderr visible rather than
reasoning about them.

**Some functions SHOULD stay undeclared.** `setTimeout`/`setInterval` are the
worked example: their IDL is `(handler, optional timeout, any... args)`, and
`setTimeout(fn)` is ordinary guest code. `FuncSig` carries a fixed
`Vec<ValType>` and the Component Model has no optional parameter, so any single
arity is wrong for some legal call. Declaring `2` would warn on correct code —
and a warning that fires on correct code teaches everyone to ignore the one that
matters. Leave those undeclared until the mechanism can express a minimum arity.

---

## Step 3 — Fix priority ladder

Work top-to-bottom. Do NOT jump levels.

```
LEVEL   WHAT                                    RISK / AUTH
─────   ────                                    ──────────
  1     Fix the test expected value              zero risk
  2     Add/fix profile entry (builtins,         zero risk — TOML only
        value_methods, intrinsics)
  3     Walker normalization (walker.rs)          no auth needed
  4     Grammar fix (grammar.pest)               no auth needed
  5     normalize_class.rs                       no auth needed
  6     Per-language emitter adapter              moderate — check regressions
        languages/<lang>/src/emitter/*_adapter.rs
        C: platforms/libc/*_adapter.rs
  7     Shared emitters (primitives/<topic>.rs)     moderate — can break other langs
  8     Compiler core (mod.rs, expressions.rs,    requires care + regression check
        calls.rs, classes.rs)
  9     host fn in platforms/{ecma,wasi,web,node,vybe}  REQUIRES explicit per-fn user auth
 10     vybe_runtime VM                          REQUIRES explicit user auth
                                                  WASM-compliant only
```

**Never reach for level 9 or 10 when levels 1–8 can solve it.**

**Level 9 is ADDING or CHANGING a host fn.** Attaching a `HostFnDecl` signature to a
host fn that already exists is not level 9 — it adds metadata, changes no behaviour,
and turns a class of silent wrong-argument bugs into compile-time warnings. Declaring
is cheap and safe; registering a new function is not.

### Layer priority — WASM first, emitter second, ecma:* third

When implementing or routing an operation:

1. **WASM opcodes** — if a WASM opcode already does what you need (arithmetic, comparisons,
   control flow, memory, GC), use it directly. Cheapest and most correct.
2. **Compiler emitters** — `primitives/<feature>.rs` or `languages/<lang>/src/emitter/<feature>_adapter.rs`.
   Compose existing opcodes and host fns. No new host fns needed.
3. **ecma:\* host functions** — if an existing `ecma:*` function maps directly, use
   `self.import("ecma:<module>", "<name>")`. ECMA-262/402/WebIDL only.
4. **Language-specific adapters** — `languages/<lang>/src/emitter/<feature>_adapter.rs` for
   runtime semantics that can't be walker-normalized. Emits bytecode composing existing fns.

---

## Common emitter pitfalls (learned the hard way)

### `emit_dyn_to_bool` steals local slots
`emit_dyn_to_bool` in `primitives/ops.rs` uses `alloc_scratch` which increments `chunk.local_count`. If called during constructor default-param checks or inside emitter helpers, the scratch slots collide with `define_local` slots allocated afterwards. **Fix:** when you already have an I32(0/1) result (from `wasm:js-undefined.test`, `REF_IS_NULL`, `I32_EQZ`), do NOT call `emit_dyn_to_bool` — use the I32 directly.

### `collect_closure_captured_in_expr` must recurse into ALL expression types
Missing: `Await`, `Spread`, `Yield`, `New`. Closures inside `await` expressions won't detect captured variables if the collection function doesn't recurse. Always check this function when adding new expression kinds.

### Promise chain methods (`.then`/`.catch`/`.finally`) must use `primitives/promises.rs`
All promise chain methods use WASM JSPI `await` via `primitives/promises.rs`. Zero host calls in the emitter. The host's `dispatch_promise_method` is a separate path for runtime fallback — the compiler should NOT emit calls to it.

### Async/await uses JSPI — not custom opcodes
`await x` → `call $jspi.await(x)`. Fulfilled promises return synchronously (correct per JSPI spec). Pending promises suspend the fiber. The VM's event loop resumes. Stack-switching (`suspend`/`resume`) is for generators only.

### Iterator protocol is pure WASM bytecode (primitives/generators.rs)
`emit_drain_custom_iterable` implements ECMA-262 §7.4.2–§7.4.5 using only WASM GC opcodes:
1. **§25.1.2 fast path**: `STRUCT_GET "next"` — if the object already has `next`, it IS the iterator (iterators are their own `[Symbol.iterator]`). Use it directly.
2. **§7.4.2 GetIterator**: `STRUCT_GET "Symbol(@@iterator)"` — TypeRegistry resolves this for built-in types (Array→`ecma:array.values`, Map→`ecma:map.entries`, Set→`ecma:set.values`, String→`ecma:string.iterator`). Fallback to walker-normalized `"iterator"` key for custom classes.
3. **§7.4.4–5 IteratorStep loop**: `STRUCT_GET "next"` → `CALL_REF 1` → check `done` → push `value`. Pure WASM block/loop/br_if.
4. **§23.1.2.1 array-like fallback**: if no `next` AND no `[Symbol.iterator]`, check `length` property and iterate by numeric index using `ARRAY_GET`. Pure WASM f64 counter loop.

**Known pitfalls:**
- ArrayIterator has `ObjectKind::Array` — the TypeRegistry resolves `Symbol(@@iterator)` to `values()` which creates an infinite chain. The §25.1.2 fast path (check `next` first) prevents this.
- Strings are opaque per `wasm:js-string` spec — `ecma:string.iterator` is a host function (same category as `charCodeAt`). This is the ONLY host call in the iterator path.
- Branch depths in the drain are tricky: the `if(no-next)` → `if(no-Symbol.iterator)` nesting means `br(0)` exits the inner if, not the outer block. When no iterator is found, `br(0)` falls through to the array-like fallback — do NOT `br(1)` to exit the outer block (that skips the fallback and returns empty).
- `emit_spread_iterable` (collections.rs) checks `isGenerator` first → drains via stack-switching. Non-generators go through `emit_drain_custom_iterable`.
- The host `iterForOf` (ecma:object) is a compatible fast path used by `emit_iter_for_of` for for-of loops. Both paths produce the same results — the host materializes eagerly while the WASM path iterates lazily via `{next() → {value, done}}`.

### Class methods need shared env for closures
`compile_lambda_direct` creates shared env arrays for captured locals, but the class method compiler in `classes.rs` originally didn't. Fixed: class methods now create shared envs when `current_closure_captured_locals` is non-empty (same mechanism). Without this, closures inside class methods (e.g. `[Symbol.iterator]()` returning `{next() {...}}`) can't access the method's locals.

Also: `compile_destructure_bind` (for `const {a, b} = obj`) must sync to the shared env after `LOCAL_SET`. Without this, destructured bindings aren't visible to closures — causes infinite loops when the closure reads NULL instead of the real value.

### Branch depths in hand-emitted bytecode
When emitting `IF`/`BLOCK`/`LOOP` manually in emitter helpers, `br(N)` exits the Nth enclosing control structure (0 = innermost). Getting N wrong is the #1 cause of "works for simple cases, hangs or returns empty for complex cases." **Debugging method:**
1. Draw the nesting stack on paper: `outer_block(2) → if(1) → inner_if(0)`
2. From your `br` position, count up to the target
3. Test with BOTH the happy path AND the fallback path — branch bugs often only trigger on the fallback

Common mistake: `br(1)` from inside `if { if { ... } }` targets the outer if, NOT the block above it. If you need to exit the block, that's `br(2)`.

### TypeRegistry `resolve_method` lowercases ALL keys
`resolve_method(type_id, name)` does `name.to_lowercase()` before HashMap lookup. Method names stored via `t.methods.insert(key, ...)` must use lowercase keys. If you insert `"Symbol(@@iterator)"` but the lookup lowercases to `"symbol(@@iterator)"`, it won't match. Always insert with the lowercase form.

### Two compilation paths for class methods vs lambdas
Class methods compile through `classes.rs` (line ~1310–1480). Lambdas compile through `compile_lambda_direct` in `calls.rs`. Both paths must handle:
- Shared env creation for captured locals
- `__js_this` capture for arrow functions
- Default parameter evaluation
- Generator entry control

When adding a new mechanism to `compile_lambda_direct`, check if `classes.rs` also needs it. The class method path is a separate code path — it does NOT call `compile_lambda_direct`.

### Isolating bugs: bisect by narrowing the reproduction
When a test fails, carve out the SMALLEST standalone program that reproduces it. Then bisect by toggling ONE variable:
- Class method vs standalone function (different compilation paths)
- With closure vs without closure (shared env creation)
- With destructuring vs without (shared env sync)
- Spread vs direct call (iterator drain vs host call)
- Custom class vs built-in type (TypeRegistry dispatch)

If it works as a standalone function but fails in a class method, the bug is in `classes.rs`, not the emitter.

### Infinite loops from closure bugs
If a test hangs (>5s), the most common causes are:
1. **Iterator drain never exits** — `done` check doesn't work, or `next` resolves to wrong function
2. **Closure reads NULL** — shared env not created or not synced, loop variable never advances
3. **Infinite recursion** — closure captures wrong binding, method calls itself instead of callback

**Quick diagnosis:** run with `timeout 3` to confirm hang, then check: does the SAME code work in a standalone function? If yes → class method compilation path. Does it work without closures? If yes → shared env.

### Host path vs WASM path compatibility
When both a host function and a WASM bytecode path exist for the same operation (e.g. `iterForOf` host vs `emit_drain_custom_iterable` bytecode), they MUST produce identical results. Test with both paths. The WASM path is the primary — the host path is the compatible fast path. If they diverge, fix the one that's wrong per spec.

---

## Step 2.5 — Profile-level fixes (often the quickest win)

Many test failures are simply **missing profile entries**. The profile TOML tells the
compiler how to emit bytecode for builtins and methods. If a builtin isn't in the profile,
the compiler treats it as a regular function call — which fails at runtime because there's
no such global.

### Where profiles live
```
languages/<lang>/src/profile
```

### Profile sections that fix tests

**`[builtins]`** — top-level function names the language recognizes:
```toml
[builtins]
print       = { emit = "print",                    min_args = 1, max_args = 255 }
parseInt    = { emit = "common:to_int",             min_args = 1, max_args = 2 }
Math.floor  = { emit = "opcode:f64_floor",          min_args = 1, max_args = 1 }
Math.abs    = { emit = "opcode:f64_abs",            min_args = 1, max_args = 1 }
Math.PI     = { emit = "intrinsic:math_pi",         min_args = 0, max_args = 0 }
```

**`[value_methods]`** — instance method calls (`obj.method(args)`):
```toml
[value_methods]
push        = { emit = "common:collections.push",  min_args = 1, max_args = 255 }
join        = { emit = "common:collections.join",   min_args = 0, max_args = 1 }
toUpperCase = { emit = "common:str_to_upper",       min_args = 0, max_args = 0 }
indexOf     = { emit = "host:ecma:string:indexOf",  min_args = 1, max_args = 2 }
```

**`[intrinsics]`** — compile-time constant data or multi-opcode templates:
```toml
[intrinsics]
math_pi     = { kind = "f64", value = "3.141592653589793" }
math_e      = { kind = "f64", value = "2.718281828459045" }
```

**`[array_methods]`** — routes HOF methods through special dispatch:
```toml
[array_methods]
map       = "__array_map"
filter    = "__array_filter"
reduce    = "__array_reduce"
sort      = "__array_sort"
```

**`[compiler]`** — semantic flags that change compilation behavior:
```toml
[compiler]
switch_fallthrough = true     # JS: cases fall through without break
hoist_var = true              # JS: var declarations are hoisted
dynamic_add = true            # JS: + overloaded for string concat
array_indexing = "one_based"  # Lua, VB, Pascal, Fortran
```

### Emit format preference order

| # | Format | Example | When |
|---|---|---|---|
| 1 | `opcode:<name>` | `opcode:f64_abs` | 1:1 WASM opcode |
| 2 | `intrinsic:<name>` | `intrinsic:math_pi` | Compile-time constant or fixed opcode sequence |
| 3 | `common:<module>.<fn>` | `common:collections.push` | Cross-language shared emitter (`primitives/<module>.rs`) |
| 4 | `common:<lang>.<fn>` | `common:php.date` | Language-specific adapter |
| 5 | `host:<module>:<fn>` | `host:ecma:date:now` | Direct ecma/wasi host call |
| 6 | `stdlib:<name>` | `stdlib:sorted` | Shared bytecode chunk |
| 7 | `invoke:<Method>` | `invoke:Peek` | Polymorphic runtime dispatch |
| 8 | `noop` | `noop` | Intentional no-op |

### How to diagnose a missing profile entry

If a test fails with `undefined is not callable` or `undefined` where a function result
is expected, check whether the function name is in the profile:
```bash
grep "functionName" languages/<lang>/src/profile
```
If it's missing, add it with the appropriate emit format. Check what JS uses for the same
concept — most things map to an existing `common:*` or `ecma:*` entry.

---

## Debugging tools — complete reference

> For the full debugging surface — static dumps, tracing, the interactive step
> debugger, VS Code/DAP, and hot reload — see
> [`documentation/debugging.md`](debugging.md). The test-focused essentials follow.

### Run a single test
```bash
cargo test -p vybe_language_js --test js -- test_module::test_name 2>&1 | grep -E "(FAILED|left:|right:|panicked)"
```

Or, without a rebuild, on the extracted file:
```bash
testrunner run tests/js/getter_setter_deep/defineProperty_creates_accessor.js
testrunner run tests/js/getter_setter_deep      # one category
testrunner run js                               # the whole suite, by name
```
`testrunner` links **no** Vybe crates — it drives the already-built `vybex`
binary — so touching the compiler does not rebuild it. Output is `cargo test`
shaped (`--progress` for a live per-suite table); `--runtime "node"` runs the
same files on the real toolchain.

### Run a standalone file
```bash
cargo run -q -p vybex -- /tmp/r.js 2>/dev/null            # run source file
cargo run -q -p vybex -- /tmp/r.js 2>/dev/null | od -c    # exact bytes
```

### Dump the common AST (diagnose walker bugs)
```bash
VYBE_DUMP_AST=1 cargo test -p vybe_language_js --test js -- test_foo::bar 2>&1 | head -200
```
Or with the CLI:
```bash
cargo run -p vybex -- --dump-ast /tmp/r.js
```
**When to use:** The AST is wrong (wrong operator, missing normalization, wrong node type).
Look for: missing fields, wrong `BinOp`, un-normalized language idioms.

### Dump bytecode (diagnose primitives/emitter bugs)
```bash
VYBE_DUMP_BC=1 cargo test -p vybe_language_js --test js -- test_foo::bar 2>&1 | head -200
```
Or with the CLI:
```bash
cargo run -p vybex -- --dump /tmp/r.js                    # all chunks
cargo run -p vybex -- --dump --chunk main /tmp/r.js       # specific chunk
```
**When to use:** The AST looks right but the wrong bytecode is emitted. Look for: wrong
opcode, wrong import index, missing setup code, wrong local slot.

### Trace VM execution (diagnose runtime bugs)
```bash
VYBE_TRACE=1 cargo test -p vybe_language_js --test js -- test_foo::bar 2>&1 | head -300
```
Or with the CLI:
```bash
cargo run -p vybex -- --trace /tmp/r.js 2>&1 | head -300
```
**When to use:** The bytecode looks right but produces wrong results at runtime. Shows
every opcode executed + stack state. Very verbose — always pipe to `head`.

### Step through a failing case interactively
```bash
cargo run -p vybex -- /tmp/r.js --debug     # pause on entry, REPL on stdin
```
Set a breakpoint (`b <line>` / `bf <fn>`), `c` to run to it, then inspect: `bt`
(call stack), `locals` (variables + values), `p <expr>` (evaluate), `set` (mutate
and continue). Faster than reading a full `--trace` firehose when you can localize
the failure to a function or line. Full command list and language support:
[`documentation/debugging.md`](debugging.md).

### Run a whole language suite (get pass/fail count)
```bash
cargo test -p vybe_language_js --test js   2>&1 | tail -3
cargo test -p vybe_language_vb --test vb   2>&1 | tail -3
cargo test -p vybe_language_php --test php  2>&1 | tail -3
cargo test -p vybe_language_c --test c    2>&1 | tail -3
cargo test -p vybe_language_dart --test dart 2>&1 | tail -3
cargo test -p vybe_language_lua --test lua  2>&1 | tail -3
```

### Save/compare baselines
```bash
# Save baseline
cargo test -p vybe_language_js --test js 2>&1 | grep FAILED | sort > /tmp/baseline.txt

# After changes, compare
cargo test -p vybe_language_js --test js 2>&1 | grep FAILED | sort > /tmp/new.txt
diff /tmp/baseline.txt /tmp/new.txt
```

### Inspect from source code (useful in helpers.rs)
```rust
// In a test helper, print the AST before compiling:
eprintln!("{:#?}", module);

// In a test helper, print chunks before running:
for (i, chunk) in chunks.iter().enumerate() {
    eprintln!("chunk {}: {} ops", i, chunk.code.len());
}
```

### Debug order (follow this sequence)
1. **Parse first** — `--dump-ast` to verify grammar + walker shape
2. **Inspect emitted chunks** — `--dump` to confirm the right emit path
3. **Trace runtime** — `--trace` only after AST and chunks look right
4. **Compare against real runtime** — `node`, `php`, `python3`, `/usr/bin/cc`

---

## Walker normalization — the most common fix layer

Each language has a `walker.rs` that parses source into the **common JS-shaped AST**.
Language-specific idioms must be rewritten here so the compiler never sees them.

### The rule: normalize in the walker, not in Rust emitters or the VM

```
PHP  is_string($x)              →  typeof $x === "string"
PHP  M_PI                       →  Math.PI
PHP  str_pad($s,$n," ",PAD_LEFT)→  s.padStart(n, " ")
PHP  count($arr)                →  arr.length
VB   Len(s)                     →  s.length
VB   Left(s, n)                 →  s.substring(0, n)
VB   UBound(arr)                →  arr.length - 1
Lua  table.insert(t, v)         →  t.push(v)     (via profile)
Lua  #t                         →  t.length       (via profile)
Python len(x)                   →  x.length
Python x.append(v)              →  x.push(v)
Fortran SIZE(arr)               →  arr.length
Dart list.length                →  list.length    (already JS-shaped)
Ruby arr.length                 →  arr.length     (already JS-shaped)
```

If a language construct maps directly to something JS/ECMA already expresses, rewrite it
in the walker. Do **not** add a language-specific intrinsic, a `__vybe_*` polyfill, or a new
host function for something the common AST + existing host fns already cover.

### Known walker pitfalls (apply to all languages)

**Pest wrapper rules** — Non-silent alternation rules produce a wrapper pair. If your
grammar has `property_name = { computed_name | ident_name }`, walkers that match on the
inner rule will silently drop the match unless they unwrap the outer pair first.

**For-loop expression ordering** — The C-style for grammar is
`init? ~ ";" ~ cond? ~ ";" ~ update?`. When init is absent, the walker sees two
`expression` children. Expression 0 is ALWAYS cond, expression 1 is ALWAYS update.
Do NOT use init-presence as a heuristic to decide which expression is which.

**Comma expressions in for-update** — `for(; i < n; i++, j--)` — the update is a
sequence expression. The walker must walk it as a single expression, not split it.

**Per-iteration bindings** — `for (const x of iter)` / `for (let i = 0; ...)` — closures
inside the body capture a fresh copy per iteration. Implement this by IIFE-wrapping the
body **only when `body_contains_closure()` is true** — otherwise IIFE breaks break/continue.

**Template literal escapes** — The grammar captures raw backslash sequences. The walker
must call `unescape_template()` on ALL text parts (full, head, middle, tail), not just
cooked parts of tagged templates.

**Computed method keys** — `{ [expr]() {} }` — detect when the key is a runtime
expression vs a well-known symbol alias (e.g. `Symbol.iterator`). Well-known symbols
become static string keys; true runtime expressions become `ObjectProperty::Computed`.

### Normalizer anti-patterns (learned the hard way)

**NEVER wrap operators/calls in IIFEs for dispatch.** The Lua normalizer (written by
Cursor) wrapped every arithmetic op, every comparison, every non-builtin call, and every
member access in immediately-invoked function expressions (IIFEs) for metamethod dispatch.
A 4-line program produced 34 chunks and 69K+ instructions. The recursion depth in
`collect_closure_captured_idents` caused stack overflows. The fix: replace every IIFE
with the simple `Binary`, `Unary`, or `Call` AST node and let the compiler handle it.
The normalizer's job is to produce the *simplest possible* JS-shaped AST, not to
implement runtime dispatch via generated lambda chains.

**Double normalization = double transform.** If a desugaring function calls
`normalize_expr(value)` on a sub-expression, and the caller (`normalize_stmt` for
`Assign`) ALSO calls `normalize_expr(value)`, the transform gets applied twice. The Lua
`try_desugar_lua_index_assign` did this — `lua_one_based_index` subtracted 1 twice
(i-1-1 = i-2), causing ipairs values to be off by one. **Rule: a desugaring helper
either normalizes its own sub-expressions OR documents that the caller must, never
both.** When moving normalize calls, put them inside the match arms that need them,
not at the top of the function.

**Normalizers produce plain AST, not bytecode strategies.** The walker doesn't know about
chunks, imports, or opcodes. It produces `ExprKind::Binary { op: BinOp::Add, left, right }`
— not a lambda that calls a metamethod lookup function. If the language has runtime
dispatch needs (Lua metamethods, Python `__add__`), those go in emitter adapters, not
in the walker as generated code.

---

## Compiler-emit pitfalls (level 7–8)

**Compile-time guard ≠ runtime guard.** A flag like "emit this setup only once" suppresses
*duplicate emission at compile time*. It does NOT stop the single emitted instruction from
*re-executing* at runtime. If that instruction sits in a loop body, it runs every iteration.
**Rule: setup that must happen once at runtime belongs at a declaration / function-entry
point, not inlined at a use-site that may be in a loop.**

**Idempotent-at-compile is not idempotent-at-run.** Same root cause: re-entrancy of emitted
code is a runtime property. Guard it with a runtime check or by hoisting, never with a
compile-time `HashSet::insert` returning false.

**Reads/writes must agree on the boxed shape.** When a binding is promoted to a cell, every
read and write must go through the cell. If some accesses are cell-aware and others read the
raw slot, you get value-vs-box mismatches.

**Per-chunk import tables.** Each chunk has its own import table. `add_import()` must be
called on the chunk that emits the `CALL_IMPORT` — calling it on `chunks[0]` when the code
runs in `chunks[current]` creates an import-index mismatch. This was the root cause of 35
Proxy failures.

---

## Per-language emitter adapters — level 6

For language-specific runtime functions that cannot be expressed as pure walker
normalization, use an **emitter adapter**:

```
languages/<lang>/src/emitter/<feature>_adapter.rs
```

Examples already in place:
- `php/emitter/string_adapter.rs` — PHP string functions → bytecode
- `php/emitter/math_adapter.rs` — PHP math functions → bytecode
- `php/emitter/array_adapter.rs` — PHP array functions → bytecode
- `php/emitter/datetime_adapter.rs` — PHP date/time functions → bytecode
- `vb/emitter/financial_adapter.rs` — VB financial functions → bytecode
- `vb/emitter/datetime_adapter.rs` — VB date/time functions → bytecode
- `python/emitter/runtime_adapter.rs` — Python builtins → bytecode
- `dart/emitter/string_adapter.rs` — Dart string methods → bytecode
- `fortran/emitter/math_adapter.rs` — Fortran math → bytecode
- `go/emitter/runtime_adapter.rs` — Go runtime → bytecode
- `cobol/emitter/*.rs` — COBOL picture formatting, data, control, files

**Rule**: adapters emit bytecode by composing existing host fns and opcodes. They do NOT
add new host functions. They do NOT call into JS polyfill files.

---

## Shared compiler emitters — level 7

`crates/vybe_compiler/src/primitives/` contains the emit modules shared across ALL languages,
alongside the AST walkers that call them. **`vybe_compiler::primitives` IS the emitter** — there
is no emitter crate and no `emitter/` module. Inside the compiler: `use crate::primitives as
common;`. From a language or platform crate: `vybe_compiler::primitives::ops`, `…::collections`, …

| Module | Operations | Profile prefix |
|---|---|---|
| `collections.rs` | push, pop, shift, slice, join, sort, reverse, indexOf, contains, concat, flat, length | `common:collections.*` |
| `dict.rs` | new, set, get, has, delete, keys, values, entries | `common:dict.*` |
| `strings.rs` | toString, concat, interpolation | (used directly by compiler) |
| `math.rs` | abs, floor, ceil, sqrt, round, min, max, pow, trig | `common:math.*` |
| `classes.rs` | object construction, methods, inheritance, type stamping | (used directly by compiler) |
| `functions.rs` | function chunks, default params, closures, async | (used directly by compiler) |
| `closures.rs` | shared env closure model — WASM GC array | (used directly by compiler) |
| `errors.rs` | try/catch/finally, exception construction, throw | (used directly by compiler) |
| `loops.rs` | for-in, HOF (map, filter, forEach, reduce) | (used directly by compiler) |
| `io.rs` | print via WASI, file I/O | `emit = "print"` |
| `convert.rs` | toInt, toFloat, toString, toBool | `common:to_int` etc. |
| `ops.rs` | dynamic equality, truthiness, typeof coercion | (used directly by compiler) |

**Use these helpers — never emit raw opcodes in compiler.rs or walkers:**
```rust
// WRONG
self.emit(Op::ARRAY_PUSH);

// RIGHT
emitter::collections::emit_push(&mut self.chunks, self.current, line);
```

Verify no raw opcodes leaked:
```bash
grep -rn "Op::ARRAY_\|Op::STR_\|Op::DICT_" crates/vybe_compiler/src/
```

---

## ecma:* host functions — what already exists

Registered by the `ecma` platform plugin (`platforms/ecma/`). Use via `self.import("ecma:<module>", "<name>")`.
**Check here before requesting new host functions — most things are already wired.**

| Namespace | Key functions |
|---|---|
| `ecma:array` | `push pop shift unshift splice slice indexOf find filter map reduce sort flat flatMap from isArray groupBy at findLast findLastIndex toSorted toReversed toSpliced with` |
| `ecma:string` | `split join slice indexOf includes startsWith endsWith replace replaceAll trim trimStart trimEnd padStart padEnd repeat at normalize matchAll fromCodePoint codePointAt` |
| `ecma:object` | `keys values entries fromEntries assign create freeze seal defineProperty getOwnPropertyDescriptor getPrototypeOf setPrototypeOf hasOwn` |
| `ecma:map` | `new get set has delete clear size keys values entries forEach` |
| `ecma:set` | `new add has delete clear size keys values entries forEach` |
| `ecma:math` | `floor ceil round abs max min pow sqrt log log2 log10 sin cos tan asin acos atan atan2 sinh cosh tanh sign trunc cbrt hypot clz32 fround imul random` |
| `ecma:json` | `stringify parse` |
| `ecma:regex` | construction, exec, match, replace |
| `ecma:date` | `new now getFullYear getMonth getDate getHours getMinutes getSeconds getTime toISOString toLocaleDateString` etc. |
| `ecma:promise` | `resolve reject all allSettled any race withResolvers` |
| `ecma:reflect` | `apply construct defineProperty deleteProperty get getPrototypeOf has ownKeys set setPrototypeOf` |
| `ecma:proxy` | `new revocable` |
| `ecma:symbol` | `new for keyFor` + well-known symbols |
| `ecma:value` | `isGenerator typeOf instanceof` |

**ecma:* is ECMA-262/402/WebIDL only.** No language-specific functions in ecma:* namespaces.

---

## C / libc runtime

C is the one frontend with true pointers and manual memory. Its runtime is a shared
**platform crate** at `platforms/libc/src/emitter/` (`vybe_platform_libc`, not a per-language
emitter). Go, Fortran, and any pointer language target it.

Key files: `stdio` adapters, `string`/`ctype`/`math` adapters, `c_runtime.rs`.
Pointer/addressable backing storage primitives live in `crates/vybe_compiler/src/primitives/`.

- **Two pointer shapes**: scalar cell `{__ref_kind:"cell", __value}` and carray
  `{__ref_kind:"carray", __base, __idx}`.
- **Integers are f64** with modular normalization via `coerce_c_value_for_type_hint`.
- **stdin** is real WASI (`wasi:cli/stdin` + `wasi:io/streams`), not a stub.
- Run the suite with `--test c` (the binary is named `c`).

**Know the representation limits.** Values are f64 objects/arrays, not bytes in linear
memory. Tests needing byte-addressable substrates can't be solved by point fixes — surface
to the user as a representation decision.

---

## Cross-language type normalization

All languages' native data structures map to the same VM types:

| Language construct | VM type | Shared emitter |
|---|---|---|
| JS `[]`, Python `list`, Ruby `Array`, Dart `List`, VB array | `ObjectKind::Array` | `collections.*` |
| JS `{}`, Lua `table`, VB object | `ObjectKind::Ordinary` | `dict.*` |
| JS `Map`, Python `dict`, PHP `array`, Ruby `Hash`, Dart `Map` | `ObjectKind::Map` | `ecma:map.*` |
| JS `Set`, Python `set`, Ruby `Set`, Dart `Set` | `ObjectKind::Set` | `ecma:set.*` |
| JS `ArrayBuffer`, Python `bytes` | `ObjectKind::ArrayBuffer` | `ecma:arraybuffer.*` |

---

## Common root-cause clusters

When many tests fail with the same pattern, look for these root causes:

| Cluster | Signal | Usual fix |
|---|---|---|
| Missing builtins | `undefined is not callable` for known function names | Add profile `[builtins]` entry |
| Missing value methods | `undefined is not callable` on `obj.method()` | Add profile `[value_methods]` entry |
| Wrong normalization | Correct structure, wrong values | Fix walker rewrite |
| Missing grammar rule | `Parse error` or panic in walker | Add grammar.pest rule |
| Closure capture bug | Functions return undefined or wrong values; `null is not callable` | Check `closures.rs`, shared env model |
| Off-by-one | Values shifted by 1 | Check array indexing (0-based vs 1-based), double normalization |
| try/catch/finally | catch runs without error, or finally doesn't run | Check compiler try/catch emit in mod.rs |
| Iterator protocol | `undefined is not callable` in for-of | Check Symbol.iterator wiring |
| `this` binding | Arrow functions pick up wrong `this` | Check compiler `this` emit, .call/.apply/.bind |
| Async/await | Empty output, hangs, or wrong ordering | Check async wrapper, JSPI await emit |

---

## .NET / Dotnet platform — shared surface for VB and C#

The dotnet platform (`platforms/dotnet/emitter/` + `emitter/dotnet/dispatch.rs`) is the
adapter layer that makes `.NET`-shaped APIs available to both VB and C#. It is NOT a host
layer — it emits pure WASM bytecode that calls existing `ecma:*` / `wasi:*` host functions.

### Architecture

```
C# source: System.Math.Sin(x)
    │
    ▼  C# walker keeps Member chain: System → Math → Sin
    │
    ▼  Compiler flatten_member_chain → ["System", "Math", "Sin"]
    │  try_compile_builtin("System.Math.Sin") — profile match if present
    │  try_compile_dotnet_component_call — component model fallback
    │
    ▼  Dotnet resolver: prefix "system.math" ∈ imports, suffix ["sin"]
    │  lookup_component_static_method("system.math", ["sin"])
    │  → ClassType "Math" in interface "dotnet.System" → MethodBody::Common("dotnet.system.math.sin")
    │
    ▼  emitter/dotnet/dispatch.rs: "dotnet.system.math.sin"
    │  → ecma:math.sin host call (pure WASM CALL_IMPORT)
    │
    ▼  VM executes standard WASM bytecode
```

### Adding a new System.* function

1. **Component descriptor** — add a `ClassType` entry in
   `platforms/dotnet/emitter/core/component_classes_system.rs`:
   ```rust
   DotnetClassExport::new(
       "dotnet.System",
       ClassType::new("Math")
           .with_method(MethodDef::static_method(
               "NewFunc", 1,
               MethodBody::Common("dotnet.system.math.newfunc".into()),
           )),
   )
   ```
   Both VB and C# get it automatically through the component model resolver.

2. **Dispatch arm** — add the emit logic in `emitter/dotnet/dispatch.rs`:
   ```rust
   "dotnet.system.math.newfunc" => {
       let idx = chunks[current].add_import("ecma:math", "newfunc");
       chunks[current].emit_call(idx, 1, line);
   }
   ```

3. **No profile entries needed** — the compiler's `try_compile_dotnet_component_call`
   resolves `System.Math.NewFunc` through the component descriptor automatically.
   Profile entries for `Math.NewFunc` (without `System.`) are still needed if you want
   the unqualified form to work — those are language-specific shorthands.

4. **Constants** (e.g. `System.Math.PI`) — cannot go through the component model's
   `MethodBody` (they are values, not calls). Use `namespace_constants` in the profile
   for the unqualified form (`Math.PI`) and walker-level constant resolution for the
   qualified form (`System.Math.PI` → literal float in the C# walker's
   `canonicalize_member_access`).

### Key files

| File | Role |
|---|---|
| `platforms/dotnet/emitter/core/component_classes_system.rs` | Component descriptor: System.Math, System.Console, System.DateTime, etc. |
| `platforms/dotnet/emitter/core/component_classes_collections.rs` | Component descriptor: Dictionary, List, Queue, Stack, HashSet, etc. |
| `emitter/dotnet/dispatch.rs` | Dispatch: `dotnet.*` → WASM bytecode |
| `platforms/dotnet/emitter/resolver.rs` | Resolves dotted names through component descriptor |
| `platforms/dotnet/emitter/core/datetime_adapter.rs` | DateTime/TimeSpan bytecode adapters |
| `platforms/dotnet/emitter/core/timespan_adapter.rs` | TimeSpan bytecode adapters |
| `primitives/calls.rs` (`try_compile_dotnet_component_call`) | Compiler fallback: flattened member chains → component model |

### Per-chunk import table — critical rule

**ALWAYS add imports to `chunks[current]`, NEVER `chunks[0]`.**

When an adapter emits `CALL_IMPORT`, the import index must be added to the same chunk
that executes the call. `chunks[0]` is the script chunk; if the code runs in a function
chunk (`chunks[current]` where `current > 0`), the import index from `chunks[0]` won't
match. This causes `null is not callable` at runtime.

```rust
// WRONG — import on chunk 0, call in current chunk
let idx = chunks[0].add_import("ecma:date", "UTC");
chunks[current].emit_call(idx, 6, line);

// RIGHT — import and call on same chunk
let idx = chunks[current].add_import("ecma:date", "UTC");
chunks[current].emit_call(idx, 6, line);
```

### .NET type mapping

| .NET type | VM type | Host surface |
|---|---|---|
| `Dictionary<K,V>` | `ObjectKind::Map` | `ecma:map.*` |
| `List<T>` | `ObjectKind::Array` | `ecma:array.*` + `collections.*` |
| `HashSet<T>` | `ObjectKind::Set` | `ecma:set.*` |
| `Queue<T>` | `ObjectKind::Array` (FIFO) | `push`/`shift` |
| `Stack<T>` | `ObjectKind::Array` (LIFO) | `push`/`pop` |
| `DateTime` | Ordinary object with Year/Month/Day/... properties | `ecma:date.*` getters |
| `TimeSpan` | Ordinary object with TotalMilliseconds/Days/... | Pure arithmetic |
| `StringBuilder` | String accumulator | `ecma:string.*` |

### Exceptions through common emitter

Use `emitter/errors.rs` for .NET exception patterns:
```rust
// Throw KeyNotFoundException when dict[key] misses
chunk.emit_string_const("The given key was not present.", line);
crate::emitter::errors::emit_exception_new_finalize(chunk, "KeyNotFoundException", line);
crate::emitter::errors::emit_throw(chunk, line);
```

---

## Regression discipline

Run the affected language suite after every change:
```bash
cargo test -p vybe_language_js --test js  2>&1 | tail -3
```

If the failure count increased, find regressions immediately — do not proceed:
```bash
cargo test -p vybe_language_js --test js 2>&1 | grep FAILED | sort > /tmp/new.txt
diff /tmp/baseline.txt /tmp/new.txt
```

Changes to `emitter/`, `primitives/mod.rs`, or `primitives/calls.rs` can break multiple
languages. But do NOT run other language suites unless the user explicitly asks.

`testrunner --json` does the baseline/diff above by itself: it writes
`results/testrunner/run_<stamp>.json` and compares it against the newest earlier
run of the same runtime, naming **regressions** and **newly-passing** tests
rather than only moving a count.

```bash
testrunner run js --json    # first run — the baseline
# … make the change …
testrunner run js --json    # REGRESSIONS (2): ✗ js/…/…
```

It is opt-in on both runs: without a saved report there is no baseline. Plain
runs write nothing, because a report per run buried `results/` in files nobody
opened.

For a text log a stats script can parse — the `<lang>.tests.txt` shape — use
`--save`:

```bash
testrunner run tests/php --save    # → results/testrunner/saved/tests.php.txt
```

One file per target, written as tests land — `tail -f` it during a long run.
Read it back with:

```bash
testrunner summary php          # failures by category, worst first
```

Only tests present in *both* runs are compared. "Absent last time" is not "was
passing last time" — without that intersection, running a whole suite after
running one category reports every other test as a regression.

Two failure modes it reports loudly rather than hiding:
- **A truncated run.** If workers cannot start (e.g. `target/` was cleaned
  mid-run) it exits 2 with `ABORTED — N of M test(s) never ran`. A partial run
  that prints a pass rate is worse than no run.
- **A hang.** The test warns once (`warning: <slug> still running after 10s`), is
  killed at the 60 s deadline and reported as `TIMEOUT` — its own verdict, not a
  `FAILED`, because it never produced an answer to be wrong. Its worker is
  replaced and the rest of the queue keeps draining, so one hang does not stall
  the run.

---

## 2026-08-13 — the extracted php suite was measured against real `php`

**Run the failing tests under `php` before believing the failure count.** Of
1352 failures, 462 were files real php also fails, and the two causes needed
opposite fixes.

⚠ **`display_errors` defaults to stdout**, so a `Warning:` or `Deprecated:`
lands *inside* `ob_start()`'s buffer and contaminates the captured value. Every
comparison here uses `php -d display_errors=stderr`. Without it the first
triage over-counted "wrong expectation" badly — `array_count_values` tests look
wrong and are not.

### 265 expectations disagreed with php — corrected from php

The program was fine; only the string in
`__vybe_check(ob_get_clean(), "…")` was wrong (e.g.
`array_map_truncates_to_shortest_input` expected `[11,22]`; php gives
`[11,22,3]` — php pads the short array with null, it does not truncate, and the
test name encoded the false belief).

The value is captured by running a COPY whose trailing check is replaced with a
marker echo, so the raw buffer is read before the harness normalises it.
⚠ Close the buffer FIRST — `echo "A", ob_get_clean(), "B"` writes `"A"` into
the still-open buffer, so the marker ends up inside the value it delimits and
the capture comes back empty.

This cannot weaken a test: the expectation now comes from php, so vybe still
fails wherever it disagrees. **196 started passing** (vybe already matched php)
and **69 still fail — real gaps that a wrong expectation had been hiding.**

### 121 files were not valid PHP at all — rebuilt from `// origin:`

`php -l` rejected them. The extraction split the origin's raw string on the two
characters `\n`, which inside a Rust **raw** string are a backslash and an n in
a php string literal, not a line break:
`printf("Name: %s, Age: %d\n", "Alice", 30)` came out as a bare line reading
`Name: %s, Age: %d`, with `"Alice"` promoted to the expectation.

Rebuilt from the origin test's own source, three shapes:

| origin shape | rebuild |
|---|---|
| `fn <name>()` / `<name> => {` holding `r#"<?php …"#` | program taken verbatim |
| a plain (escaped) Rust string holding `<?php …` | unescaped, then verbatim |
| several `assert_output("<expr>", …)` / `assert_int` / `assert_bool` | one `echo <expr>, "\n";` per assertion, so every assertion stays observable |

⚠ **Namespace form dictates where the harness can go.** With braced
`namespace X { }` php allows **no** code outside a block, so the harness gets
its own `namespace { }` and any top-level run of the program is wrapped in one
— several origin programs mix the two and were therefore never valid php.
With unbraced `namespace X;` the declaration must be the first statement, so
the preamble goes immediately after it.

Nothing is written unless php runs the rebuilt file twice with identical output
and a zero exit, so a non-deterministic or intentionally-fatal program is
reported rather than guessed at. **86 rebuilt; 35 remain**, almost all programs
whose origin expectation does not hold under php 8.4 — e.g.
`spread_with_unpacking_non_array_throws` expects `[...1]` to be catchable, and
in php 8.4 it is an uncatchable fatal.

### Result

```
1354 → 1090 failed        (class identity + new static(): −22
                           expectations corrected:        −193
                           corrupt files rebuilt:          −53)
```

⚠ Every "new failure" in each round passed when re-run **solo, twice** —
`php/namespaces`, `php/classes`, `oop`, `method_chaining` and
`intersection_types` flake under parallel load. Two back-to-back full runs with
no code change between them differed by 14 names.

Scripts: `fix_expectations.py`, `regen_from_origin.py` (session scratchpad).
