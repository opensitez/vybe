# Debugging Vybe

Everything available for inspecting, tracing, and debugging Vybe programs — from
static bytecode/AST dumps through a full interactive step debugger, VS Code
integration, and Dart-style hot reload.

Because Vybe compiles every language to the same bytecode VM, **the interactive
debugger is language-agnostic**: `vybex <file> --debug` works the same whether the
file is JavaScript, Python, PHP, Go, C, Ruby, or any other supported language. The
one feature with a per-language boundary — evaluating a *typed-in expression* — is
called out explicitly in [Language support](#language-support).

- [Quick reference](#quick-reference)
- [Getting a debuggable file out of a failing test](#getting-a-debuggable-file-out-of-a-failing-test)
- [Static inspection (no run)](#static-inspection-no-run)
- [Execution tracing](#execution-tracing)
- [The step debugger (`--debug`)](#the-step-debugger---debug)
- [VS Code / Debug Adapter Protocol (`--dap-port`)](#vs-code--debug-adapter-protocol---dap-port)
- [Hot reload (`--watch` and `reload`)](#hot-reload---watch-and-reload)
- [Language support](#language-support)
- [Environment variables](#environment-variables)
- [How it stays zero-cost when off](#how-it-stays-zero-cost-when-off)

---

## Quick reference

| Task | Command |
|---|---|
| Disassemble bytecode, don't run | `vybex file.js --dump` (`-d`) |
| Print the parsed common AST | `vybex file.js --dump-ast` |
| Trace every opcode as it runs | `vybex file.js --trace` (`-t`) or `VYBE_TRACE=1 vybex file.js` |
| Interactive step debugger (REPL) | `vybex file.js --debug` (`-g`) |
| Attach VS Code | `vybex file.js --dap-port 4711` |
| Re-run on every file change | `vybex file.js --watch` (`-W`) |
| Limit dump/trace/AST to one function | add `--chunk <name-or-index>` |
| Screenshot a GUI program, no window | `vybex app.c --capture out.png` |
| …just one control | add `--capture-control <name>` |

`--debug`, `--dap-port`, and `--watch` compose: e.g. `--watch --debug` restarts a
fresh debug session on every save.

---

## Getting a debuggable file out of a failing test

Every tool on this page takes a **file**. A test inside `cargo test` is not a
file, which is why none of them apply to one: `cargo test` cannot give you a step
debugger, a DAP server, an AST dump or a bytecode trace, and the program source
only exists as a Rust string literal inside a `#[test]`.

`crates/testrunner` exists to remove that gap. It extracts the native Rust suites
into standalone source files — one test per file, under
`tests/<language>/<category>/<name>.<ext>` — that are ordinary `.js` / `.py` /
`.go` / `.php` programs. Every command in this document then works on them
unchanged:

```sh
vybex --dump-ast tests/go/json_marshal/marshal_bool_true.go   # parse + walker shape
vybex -d         tests/go/json_marshal/marshal_bool_true.go   # emitted bytecode
vybex -t         tests/go/json_marshal/marshal_bool_true.go   # opcode trace
vybex -g         tests/go/json_marshal/marshal_bool_true.go   # step debugger
vybex --dap-port 4711 tests/go/json_marshal/marshal_bool_true.go
```

They also run on the language's real toolchain (`go run`, `node`, `python3`,
`php`, `wasmtime`), which settles "is the expectation even right?" without
writing a second repro by hand.

**You no longer copy a case into `/tmp` to step through it.** That was the old
loop — find the failing `#[test]`, retype or paste its source into `/tmp/r.js`,
hope the transcription was faithful, then debug the copy. The extracted file *is*
the program the suite ran, so `vybex -g` on it debugs the real thing.

To find that file, run the tests and read the slug — the slug **is** the path:

```
test go/json_marshal/marshal_bool_true ... FAILED
     └─ tests/go/json_marshal/marshal_bool_true.go
```

The tree is the granularity, and a target may be a suite, a category, or one
file. **Run a category, not a suite** — a suite is thousands of tests and minutes
of wall clock, a category is seconds, and a category is almost always the unit
your change touches. More than one target is fine, including across languages:

```sh
testrunner run go                                          # whole suite (slow)
testrunner run tests/go/json_marshal                       # one category
testrunner run tests/go/json_marshal tests/go/bufio_io     # two categories
testrunner run tests/go/json_marshal/marshal_bool_true.go  # one test
```

Output is `cargo test` shaped by default (`--progress` for a live table), so the
failing slugs grep straight out. Four states, four words:

```
test go/json_marshal/marshal_bool_true ... ok
test go/json_marshal/marshal_map_order  ... FAILED
warning: go/loops/infinite_range still running after 10s
test go/loops/infinite_range            ... TIMEOUT
```

`TIMEOUT` is its own verdict because a hang is not a failure — it never produced
an answer to be wrong. The `warning:` is not a verdict at all: the test is still
running and may still pass. It fires **once**, at 10 s.

On a terminal `ok` is green, `FAILED` red, `TIMEOUT` orange, `warning:` yellow;
piped to a file there are no escapes at all, so a log stays diffable.

It is **slower than `cargo test`** on a full suite, and it does not hang: a
non-terminating test is killed at the 60 s deadline (`--timeout`) and reported
while everything else keeps running.

Two things worth knowing before you trust what you see:

- **The verdict is the exit code**, not stdout. Nothing parses program output, so
  `--trace` and `-d` do not change a verdict and can be left on.
- **Warm workers are the default.** `--cold` forces a fresh process per test; if
  warm and cold disagree on a verdict, the VM reset is leaking state and that is
  the bug, not whatever you were chasing.

See [`documentation/testrunner.md`](testrunner.md) for the full tool.

---

## Static inspection (no run)

These compile the program and print information without executing it.

### `--dump` / `-d` — bytecode disassembly

```bash
vybex file.js --dump
```

Prints each compiled chunk (one per function/script) as human-readable
instructions with offsets. Constant-pool operands are **resolved inline** — a
`const`, `global.get`/`global.set`, and `struct.get`/`struct.set` show the actual
value or property/global name in parentheses, so property and global access is
readable without cross-referencing the constant pool:

```
── Chunk 4 ── whoami
0000  global.get 0 (C)
0007  struct.get 1 (prototype)
0013  struct.get 2 (__proto__)
0019  struct.get 3 (whoami)
0025  local.set 0
...
```

(`local.get`/`local.set` show a raw *slot* index — they don't index the constant
pool.) The live debugger's `dis` prints the same, and additionally resolves
host-function names for `CallHost` (the standalone `--dump` shows the raw import
index).

Use `--chunk <name-or-index>` to dump a single function:

```bash
vybex file.js --dump --chunk add
vybex file.js --dump --chunk 4
```

### `--dump-ast` — the common AST

```bash
vybex file.js --dump-ast
```

Parses the source into the **common AST** (`vybe_ast::Module`) — the
language-independent representation every frontend targets — and prints a
top-level outline. For the full recursive AST:

```bash
VYBEX_DUMP_AST_FULL=1 vybex file.js --dump-ast
```

This is the fastest way to see how a frontend lowered a construct (spans, node
kinds, desugaring) and to debug a new language frontend.

`--chunk` narrows it to one declaration, the same way it narrows `--dump` and
`--trace`:

```bash
vybex file.c --dump-ast --chunk via_helper
```

Matches top-level functions, classes, structs, interfaces, enums and namespaces
by name (case-insensitively, so it works for the case-folding languages), and
always prints them in full — no outline, no `VYBEX_DUMP_AST_FULL`. On a real
program the difference is ~20 lines instead of tens of thousands, which is what
makes "how did the walker lower this one function?" a practical question. An
unmatched name exits non-zero.

### `--emit-wasm` / `-w` — write a `.wasm` binary

```bash
vybex file.js --emit-wasm    # writes file.wasm
```

Not a debugger, but useful for inspecting the compiled module with external WASM
tooling.

---

## Execution tracing

### `--trace` / `-t` (or `VYBE_TRACE=1`)

Prints every instruction as it executes, with the current stack top:

```bash
vybex file.js --trace
VYBE_TRACE=1 vybex file.js          # same thing
```

```
  TRACE          add @0000 LOCAL_GET  stack: [10] (depth=1)
  TRACE          add @0006 LOCAL_GET  stack: [32] (depth=2)
  TRACE          add @0012 I32_ADD    stack: [42] (depth=1)
```

Filter to one chunk to cut the noise:

```bash
VYBE_TRACE=1 VYBE_TRACE_CHUNK=add vybex file.js
vybex file.js --trace --chunk add
```

`--trace` is the low-tech firehose. For an interactive, filterable version of the
same information, use the step debugger's live opcode stream (`trace on|off`,
below).

---

## The step debugger (`--debug`)

```bash
vybex file.js --debug          # or -g
```

Starts the program **paused on entry** with an interactive REPL on stdin. Type
`h` for help. The VM runs on the main thread; the REPL runs on worker threads, so
inspection and control happen while execution is genuinely suspended (blocked, not
polling).

#### Scripting the debugger (non-interactive)

Because it reads commands from stdin, you can drive a whole session from a pipe —
no TTY, no timing tricks:

```bash
printf 'bf myFunc\nc\nbt\np w\np w.field == 3\nc\nq\n' | vybex file.dart --debug
```

A resuming command (`c`, `s`, `n`, `o`, `rt`) **blocks the caller until the VM
actually stops again** — its reply is deferred to the next pause. That keeps a
piped command stream in lock-step with execution: each line is consumed only while
paused, so the next command can never race ahead into a still-running VM. **No
`sleep`s are needed** (and they never worked reliably — earlier guidance to space
commands with `sleep` was a workaround for a since-fixed desync). A `c` that runs
to program completion simply returns when the program exits.

### Command reference

Type `h` in the REPL for the built-in cheat sheet. Full list:

**Execution control**

| Command | Aliases | Action |
|---|---|---|
| `c` | `continue`, `cont` | Resume |
| `s` | `step` | Step into |
| `n` | `next` | Step over |
| `o` | `out`, `fin`, `finish` | Step out (run to return) |
| `si` | `stepi` | Single bytecode instruction |
| `rt <line>` | `runto`, `tbreak` | Run to a line once (run-to-cursor) |
| `restart` | `R` | Restart the whole program (fresh state) |
| `detach` | | Detach the debugger, let it run free |
| `skip-system on\|off` | `sys` | Skip the injected runtime prelude on entry (default **on**) — see below |
| `q` | `quit`, `exit` | Terminate |

**Breakpoints**

| Command | Action |
|---|---|
| `b <line>` | Break on a source line (slides to the nearest real code) |
| `b <file>:<line>` | Same, file-qualified |
| `b <chunk>@<offset>` | Break at a bytecode offset in a chunk |
| `b <loc> if <expr>` | Conditional breakpoint (see [expression eval](#language-support)) |
| `bf <fn> [if <expr>]` | Break on entry to a function by name — breaks on **every** chunk with that name, so overridden methods across a class hierarchy (e.g. three `whoami` chunks for `A`/`B`/`C`) all get the breakpoint (reported as `whoami (×3)`) |
| `lp <line> <msg>` | Logpoint — log `<msg>` (with `{expr}` interpolation) and keep running |
| `ignore <id> <n>` | Skip a breakpoint's first `n` hits |
| `catch throw` / `uncaught` / `off` | Break when an exception is thrown / only when uncaught / disable |
| `bl` (`breaks`) | List breakpoints |
| `bd <id>` (`delete`) | Delete a breakpoint |
| `enable <id>` / `disable <id>` | Toggle a breakpoint |

`<chunk>` may be a function name or a numeric index (see `chunks`). A line with no
directly-attributed instruction slides to the nearest one in the enclosing
function, gdb-style — so any line is breakable.

**Data watchpoints**

| Command | Action |
|---|---|
| `wp <name>` | Watchpoint — pause whenever the value changes |
| `wps` | List watchpoints |
| `unwp` (`clearwp`) | Clear watchpoints |

**Inspection**

| Command | Action |
|---|---|
| `bt` (`where`, `backtrace`) | Call stack, with a source line per frame |
| `locals [frame]` (`l`) | Local variables of a frame (0 = innermost), by name + value |
| `stack` | Operand stack of the current frame |
| `g [prefix]` (`globals`) | Globals, optionally filtered by name prefix |
| `dis [n]` (`disasm`) | Disassemble ±`n` instructions around the current ip |
| `chunks` | List all chunks (index, name, arity, size) |
| `fibers` (`threads`) | List the current + suspended fibers |

**Variables and evaluation**

| Command | Action |
|---|---|
| `p <name>[.field][idx]` | Read a local/global by name, with `.field` / `[index]` drill-down |
| `p <expr>` | Evaluate a compound expression (e.g. `p a + b`) — see [Language support](#language-support). Also inspects objects and prototype chains: `p typeof C.prototype.__proto__.whoami`, `p B.prototype.hasOwnProperty("m")`, `p x === y`. Reads see live heap state (objects are shared by reference). **Host-function calls hit the live program's host state** — e.g. calling a `vybe:gui.*` function reads/uses the running GUI, not a fresh one. Calling a *user* function from within `p` is not supported (it runs in an isolated mini-VM, which is what keeps a throwing eval from corrupting the paused program). |
| `set <name> = <literal>` | Write a literal (number / string / bool / null) into a local or global |
| `watch <expr>` | Add a watch expression, re-evaluated and shown on every pause |
| `watches` | List watch expressions |
| `unwatch` | Clear watch expressions |

**GUI debugging** (for programs that build a `vybe:gui` UI — Dart/Flutter, WinForms, VCL, …)

All GUI frameworks Vybe targets — Flutter, .NET WinForms, Pascal VCL — drive the
**same** `vybe_widgets` runtime and the **same** `GuiState`. So these commands are
framework-agnostic: nothing here is language-specific.

| Command | Action |
|---|---|
| `widgets` (`controls`) | Dump the live GUI state — every realized control, its **laid-out rect**, its properties, and its wired events. Reads the shared `GuiState` directly, so it reflects the running program even with no window open. A control showing `rect=0x0 ← never laid out` will not render and cannot be hit-tested; with no window at all, the dump says so once rather than blanking every row |
| `html [control]` (`dom`) | Serialise the live document as **HTML** — real tags, attributes and inline style, indented. `widgets` reports each control's *properties*; this reports the **structure**. With a control name, just that element and its subtree (`outerHTML`) |
| `css <control> [property] [value]` | The element's declarations. No property lists them all; one property reads it; a value **sets** it, live |
| `attr <control> <name> [value]` | Read or set a content attribute |
| `text <control> [value]` | Read or set the element's text content |
| `draws [control] [n]` (`drawlist`) | List the draw commands a canvas recorded — op, coordinates and colour, in order. Covers both real `Canvas` controls and overlay canvases. `n` caps the listing (default: all) |
| `click <control>` (`tap`) | Simulate a click: invoke the control's `Click` handler **through the live VM**. State changes take effect (e.g. Flutter `setState`, or a WinForms handler writing a control's text). Fully headless |
| `fire <control> <event>` | Simulate any event (`fire btn7 Click`, `fire form Load`) |
| `close [control]` | Fire a `Close` event on the form (or a named control) |
| `capture [control] [file.png]` | Render the live frame to a PNG. Both arguments optional — no arguments writes the whole form to `vybe-capture.png`; a lone `*.png` argument is the file, anything else is a control name |

Control names come from `widgets`. Handler arguments are built by the handler's
arity, the same rule the real window uses: `0 → ()`, `1 → (form)`, `2 → (form,
senderName)`. That fits both Flutter's 0-arg `onPressed` closures and .NET/VB
`(sender, e)` handlers with no special-casing.

Two things worth knowing:

- **Receiver.** Instance-method handlers (C#/VB `void OnClick(sender, e)`) receive
  the *form* as their receiver, and the form isn't registered until the app
  starts (`Application.Run` / `showForm`). So fire those **after** the app is
  running — e.g. break on the `Application.Run(...)` line, `n` (step over it, which
  registers the form), then `click`. Flutter's receiver-less closures fire at any
  pause.
- **Display refresh.** `widgets` reflects a handler's changes **immediately** — it
  reads `GuiState`'s property store, which `set_property` updates on every write.
  So after `click btn7` you'll see the display control go to `7` in the next
  `widgets` dump, headless, no window needed. In a live window, non-input state
  changes (a fired handler, a timer, an async update) now repaint on a ~60 Hz
  cadence, so they appear on screen without you having to move the mouse.

**The structure, and the inspector: `html` / `css` / `attr` / `text`**

`widgets` answers *what are this control's properties*. `html` answers *what is
inside what*, which is the other half of a rendering bug — a control with
perfectly correct properties in the wrong parent reads as fine in a property
dump:

```
(vdbg) html
<body>
  <div bevelouter style="dock: top; height: 72">
    <input type="text" readonly="true" style="height: 42; left: 8; top: 8; width: 296">
    <button style="height: 20; left: 8; top: 50; width: 146">C</button>
  </div>
  <button style="height: 56; left: 8; top: 88; width: 70">7</button>
</body>
```

It is also the only form of GUI evidence that **diffs**. A capture tells you
something moved; this tells you which element, and a golden file can be reviewed
in a patch. And because every frontend's goal is to *become HTML*, output that
isn't markup you'd have written by hand is itself the bug report — a `<div>`
where `<header>` belongs shows up here and is invisible in a screenshot.

The inspector reads and writes one element live. Writes go through the same
`Document` entry points a guest program uses, so setting something here does
exactly what the program setting it would do — and it lands in `widgets`, `html`
and `capture` at once, because there is only one tree:

```
(vdbg) css n2                     # every declaration
  height: 42
  left: 8
(vdbg) css n2 left                # one property
  left: 8  ← declared, computed 4px
(vdbg) css n2 left 100px          # …and set it
(vdbg) text n7                    # read the caption
  "7"
```

**`← declared, computed` is the line to look for.** Geometry is read off the
*control*, and a laid-out or docked element does not sit where its own `left`
asked — so the two figures disagreeing is a container overriding its child. A
child of a flow container reports `left: 8 ← declared, computed 4px`, while a
sibling parented to the form reports a bare `left: 82`. That difference is the
whole diagnosis, on one line.

**What was drawn: `draws`**

`capture` shows you the pixels. `draws` shows you the *commands* — and the two
answer different questions. On screen, these three failures are identical
(nothing visible), and only the command list tells them apart:

| Symptom in `draws` | Diagnosis |
|---|---|
| no commands at all | the draw call never reached the canvas — a dispatch or wiring bug |
| commands present, colour matches the background | drawn invisibly (e.g. black on black) |
| commands present on the *wrong* canvas name | the surface handle was lost — the draw landed on another control |

```
(vdbg) draws 10
  460 command(s) on `Vybe SDL Adapter - Signal Monitor_surface`
     0  setFillColor    #181c26
     1  fillRect        0,0 800x480
     2  setFillColor    #e8eef8
     3  fillText        "Signal Monitor" @24,18
  … 450 more (pass a count to see more)
```

Colours are printed as hex, so "why is everything blue?" is answerable by
reading rather than by guessing. `drawImage` prints its source dimensions, never
its pixels.

**Seeing the pixels: `--capture`**

`widgets` tells you a control's *properties*. It cannot tell you the thing that
actually matters for a rendering bug — what got drawn. `--capture` renders one
frame into an offscreen pixmap and writes it as a PNG:

```sh
vybex examples/sdl/hellosdl.c --capture frame.png
# [vybex] captured 800x600 → frame.png

vybex examples/sdl/hellosdl.c --capture chart.png --capture-control mychart
# [vybex] captured 800x480 → chart.png
```

**No window is opened and no event loop runs.** Everything a GUI program drew
during its run is already in `GuiState`, so the capture replays exactly what the
window would have shown — `FormApp::render` and the capture path call the same
`gui_capture::render_into`, so the two cannot drift apart.

Three consequences worth the flag's existence:

- A rendering bug becomes a *file* you can look at, instead of a screenshot you
  have to take by hand.
- It is non-interactive, so it works over SSH, in CI, and from a script.
- It is deterministic, so two PNGs can be diffed as a regression check.

`--capture-control` and `draws` both accept an **unambiguous substring**, matched
case-insensitively, because exact names are often unusable — an SDL surface is
named after the window *title*, so its real name may be
`vybe sdl adapter - signal monitor_surface`. All of these hit it:

```sh
--capture-control surface      # substring
--capture-control MONITOR      # any case
(vdbg) draws monitor
```

A substring matching more than one control is rejected rather than guessed.

`--capture-control` crops to `Form::get_control_rect`. An **overlay canvas** — an
SDL surface, say — is painted onto the form rather than being a child control, so
it has no rect of its own and falls back to the form's. A name that matches
nothing lists the names that do exist and **exits non-zero**, so a broken capture
in a script fails loudly rather than leaving a stale file:

```
[vybex] capture failed: no control named `chart` (have: canvas1, monitor_surface)
```

The debugger's `capture` is the same renderer at a **breakpoint**, which makes it
strictly more powerful: it photographs a *partial* frame. Break inside a drawing
routine and you see exactly how far the program got —

```bash
printf 'bf plot_wave\nc\nc\ncapture partial.png\nq\n' | vybex app.c --debug
```

— which is how you tell "drew nothing" apart from "drew it in the wrong place"
or "drew it, then painted over it". Without a window nothing has called
`on_init`, so the capture lays the form out to `GuiState`'s size first;
otherwise every control would still have a zero rect and the frame would come
out blank.

**Hot reload & tracing**

| Command | Action |
|---|---|
| `reload` | Recompile the source and swap changed function bodies in place (see [hot reload](#hot-reload---watch-and-reload)) |
| `trace on` / `trace off` | Toggle a live, filterable opcode event stream (the interactive `VYBE_TRACE`) |
| `trace canvas on` / `off` | Toggle canvas draw-routing tracing (the interactive `VYBE_DBG_CANVAS`). Prints the requested control name and the one it resolved to for every draw — which is how a draw landing on the wrong canvas gets caught. Settable at a breakpoint, so you can leave it off until you reach the interesting frame |

### Example session

```
$ vybex add.js --debug
── vybe step debugger ── type `h` for help. Paused on entry.
(vdbg) b add:2
  breakpoint #1 set at add line 2
(vdbg) c

■ paused (breakpoint #1) — add@0025 line 2
  2 frame(s), in add
(vdbg) bt
  #0 add@0025 line 2
  #1 <script>@216343 line 7
(vdbg) locals
  a [0] = 10
  b [1] = 32
  s [2] = null
(vdbg) p a + b
  42
(vdbg) set a = 100
  a: 10 → 100
(vdbg) c
132
● program exited → null
```

`set a = 100` mutates the live frame — continuing produces `132` (`100 + 32`)
instead of `42`, confirming the write flowed into real execution.

### Skipping the runtime prelude (system code)

Every program is compiled with a language runtime **prelude** prepended to the
`<script>` chunk — for JavaScript that's ~200k instructions of intrinsics
(`Map`/`Set`, prototype wiring, `Error` hierarchy, …). By default the debugger
**skips over it**: the first pause lands on the first line of *your* code, not
deep in the prelude, and stepping never wanders into runtime internals.

This is driven by a compiler-emitted boundary marker (`chunk.user_code_offset`),
so it's exact, not a heuristic, and it's a generic mechanism (any frontend that
injects a prelude can mark the boundary).

You can still debug the prelude when you want to:

- **Set an explicit breakpoint inside it** — breakpoints in system code fire
  normally, before the auto-skip's user-code stop (e.g. `bf <prelude-fn>` or a
  `b <line>` that lands in the prelude).
- **`skip-system off`** turns the auto-skip off entirely so entry and stepping
  behave like plain bytecode.

Languages without an injected prelude are unaffected — with no marker, the entry
pause is at the first instruction as usual.

---

## VS Code / Debug Adapter Protocol (`--dap-port`)

```bash
vybex file.js --dap-port 4711
```

Starts a **Debug Adapter Protocol** server on `127.0.0.1:4711`. Attach VS Code
with a launch configuration that points at the running server:

```jsonc
// .vscode/launch.json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Attach to Vybe",
      "type": "debugpy",       // any DAP-capable type works as a generic attach
      "request": "attach",
      "debugServer": 4711
    }
  ]
}
```

The adapter maps DAP onto the same debugger backend, so you get VS Code's native
breakpoints (including **conditional** ones), stepping buttons, call stack,
variables pane, watch, **evaluate** (debug console), **set-variable**, and
restart. It emits `stopped` / `continued` / `terminated` / `output` events.

Program stdout goes to the terminal where `vybex` runs (the DAP channel is a
separate TCP connection), so program output and the debug protocol never collide.

Notes:
- The program pauses on entry until the client sends `configurationDone` /
  `continue`.
- `--dap-port` and `--debug` are mutually exclusive per run (DAP replaces the
  stdin REPL).

---

## Hot reload (`--watch` and `reload`)

Two flavors, from clean-slate to Dart-style stateful.

### `--watch` / `-W` — re-run on change (Phase 1)

```bash
vybex file.js --watch
```

Runs the program, then re-runs it (in a fresh process) on every source change.
Clean-slate: program state is *not* preserved — this is the "save → see the new
output" dev loop. Works for scripts (re-run on change) and long-running servers
(kill + restart on change), and composes with `--debug` (each reload starts a
fresh debug session).

### `reload` — stateful hot reload (Dart-style)

From inside a `--debug` session:

```
(vdbg) reload
  reloaded 1 function(s): double · 17 unchanged (heap/globals preserved)
```

Recompiles the source and **swaps only the changed function bodies in place**,
preserving the heap, globals, and the current call stack. `main` is not re-run;
the next call of a reloaded function uses the new code. You can even reload a
function you're **currently paused inside** — the running activation finishes on
the old body, and the next call uses the new one.

For safety it rejects (with a clear reason, leaving state untouched) anything that
isn't a safe body swap:
- structural changes (a function added / removed / renamed),
- a function currently live in a *suspended async task*,
- exception-tag changes.

One known limit: editing a string *literal* changes the pooled string-constant
imports on the top-level `<script>` chunk (always live), so those edits report
"restart needed" rather than reloading. Code/logic edits to functions reload
cleanly.

---

## Language support

The debugger is built on the VM, so the **structural surface works for every
language** Vybe compiles. The only feature with a per-language boundary is
evaluating a *typed-in expression*.

### Works for all languages (VM-level)

Step debugging (all breakpoint kinds, stepping, run-to-cursor, exception
breakpoints, hit-counts, logpoints), hot reload (`--watch` and `reload`), variable
inspection (backtrace, `locals`, `globals`, `stack`, `dis`, **`p <name>`**,
`set <name> = <literal>`, data watchpoints, `fibers`), the REPL, VS Code / DAP,
the live opcode stream, and restart.

> Note: for **WAST**, structural debugging works but variable *names* are
> unreliable — inspect by slot/index there.

### Typed-in expression evaluation

`p <expr>` (compound expressions), conditional breakpoints (`b <loc> if <expr>`),
and watch-*expressions* (`watch <expr>`) compile the expression with the real
compiler in an isolated mini-VM. This requires parsing an expression in the source
language, which is supported for:

- **Verified:** JavaScript/TypeScript, Python, PHP, Lua, Go, C, Java, C#, Dart.
- **Not yet:** VB, Pascal, Fortran, COBOL, WAST — these report a clear
  "use `p <name>`" message. `p <name>` reads and everything else still work.

When a structural read fails for a real reason (e.g. `p w.__control_type` on an
object that has no such field), the debugger now surfaces that actionable error
(`no field 'w.__control_type'`) rather than a generic "eval not available" —
even in languages outside the eval set.

Adding a language to the eval set is a small, localized change (a
[frontend expression-fragment adaptation](add_vybex_language.md#expression-eval-for-the-debugger));
it does not affect the compiler's normal path or non-debug runs.

---

## Environment variables

| Variable | Effect |
|---|---|
| `VYBE_TRACE=1` | Trace every opcode (same as `--trace`) |
| `VYBE_TRACE_CHUNK=<name>` | Restrict trace output to one chunk |
| `VYBEX_DUMP_AST_FULL=1` | With `--dump-ast`, print the full recursive AST |
| `VYBE_DEBUG_IMPORTS=1` | Trace per-chunk import registration/resolution |
| `VYBE_PRELUDE_DEBUG=1` | Debug language-prelude expansion (see the Python prelude notes) |
| `VYBE_DEBUG_AC=1` | Diagnostics for the async/continuation machinery |
| `VYBE_GUI_TRACE=1` | Trace GUI host calls |
| `VYBE_DBG_CANVAS=1` | Trace canvas draws and which control each one resolves to. Seeds the toggle the debugger's `trace canvas on/off` flips, so either can drive it. Pairs with `--capture`: the PNG shows *what* was drawn, this shows *where it went* — together they catch a draw landing on the wrong canvas |

Prefer these and the interactive debugger over adding `print`/`eprintln`
statements — they're purpose-built and don't pollute the source.

---

## How it stays zero-cost when off

Every interactive-debug feature is gated behind a single `instrumented` flag on
the VM (`instrumented = trace || debugger.is_some()`), checked once per
instruction at one place in the dispatch loop. In a normal `vybex file.js` run
there is no attached debugger, so `instrumented` is false and the hot path is
**byte-identical** to a non-debug build — `read_byte`/`get_constant` and the
dispatch loop are untouched.

- **Breakpoints, stepping, watchpoints, the opcode stream** live inside that one
  gated hook.
- **Expression evaluation** runs in an *isolated mini-VM*, so a throwing eval can
  never corrupt the paused program's stack, frames, or exception handlers; it's
  only reachable through the debugger's eval hook (`None` when not debugging). The
  mini-VM *shares the live program's host-function closures* (matched by name), so
  host calls in `p <expr>` reach live host state (e.g. the GUI) while execution
  stays isolated. **Event simulation** (`click`/`fire`) instead runs the handler
  in the *live* VM via the event-fire hook, so real state changes and breakpoints
  apply.
- **Stateful `reload`** swaps chunk bodies; where it must let an old activation
  finish on old code, it *relocates* the frame to a copy rather than adding any
  per-instruction check — again, zero cost when not reloading.

So you can leave all of this in the shipping VM with no performance penalty for
programs that aren't being debugged.
