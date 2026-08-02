# testrunner

Extracts the native Rust test suites into standalone source files, and runs them
through the real `vybex` binary — or through `go run`, `node`, `python3`.

Same shape as `ecma/testecma` (warm worker pool, timestamped JSON reports,
run-over-run comparison), generalised from one language to sixteen.

## Why it exists

Two costs, both removed:

- **Rebuilds.** `languages/*/tests` is ~6,000 Rust files / 1.18M lines. Touching
  one compiler line recompiled all of it before a single test ran. Extracted
  tests are `.go` / `.py` / `.php` files — nothing to recompile. This crate
  links nothing from Vybe, so it never rebuilds either.
- **No instrumentation.** `cargo test` cannot give you the step debugger. A
  standalone file can: `vybex -g tests/go/json_marshal/marshal_bool_true.go`.

And one capability gained: the same file runs on the *real* language runtime, so
an expectation written against Vybe's behaviour can be checked against Go's.

## Build

From the repo root:

```sh
cargo build --bin vybex          # the binary under test
cargo build --bin testrunner     # → target/debug/testrunner
```

It is a normal workspace member, but with **zero Vybe dependencies** — that is
what keeps it out of the 20-crate relink, not any workspace trickery. Measured:
after touching `crates/vybe_compiler/src/bundle.rs`, `cargo build --bin
testrunner` finishes in **0.70 s** with nothing recompiled.

## Extract

```sh
# one module
./target/debug/testrunner extract \
    languages/go/tests/go/test_json_marshal.rs

# a whole language
./target/debug/testrunner extract languages/go/tests/go/*.rs
```

Writes `tests/<language>/<category>/<case>.<ext>` — one test per file, category
from the Rust module name. Extraction is additive: the `.rs` suites are never
touched.

Every case is round-tripped (decode → re-escape → compare) before it is written,
because Go and C sources arrive as escaped Rust literals with backtick struct
tags nested inside them, and a bad unescape corrupts a program silently.

Cases whose printed lines can't be paired to an expectation 1:1 — loops,
`Printf`, mismatched counts — are **reported, not guessed**, and emitted without
assertions.

## Run

```sh
# a directory, warm workers (default)
./target/debug/testrunner run tests/go

# one file — this is the `cargo test <name>` replacement
./target/debug/testrunner run \
    tests/go/json_marshal/marshal_bool_true.go

# against the real Go toolchain instead of vybex
./target/debug/testrunner run tests/go --runtime "go run"

# a fresh process per test — use to prove warm mode isn't leaking state
./target/debug/testrunner run tests/go --cold
```

```sh
# a live per-suite table instead of cargo-shaped output
./target/debug/testrunner run go python wast --progress
```

Output is `cargo test` shaped by default — `test <slug> ... ok` per line, a
failures block, then a `test result:` summary — so it greps and diffs. Four
states, four words:

```
test go/json_marshal/marshal_bool_true ... ok
test go/json_marshal/marshal_map_order  ... FAILED
warning: go/loops/infinite_range still running after 10s
test go/loops/infinite_range            ... TIMEOUT
```

`TIMEOUT` is its own verdict: a hang never produced an answer to be wrong. The
`warning:` is not a verdict — the test is still running and may still pass — and
fires **once**, at 10 s. The deadline is **60 s** (`--timeout`), after which the
worker is killed and replaced while the others keep draining.

A target may be a whole suite (`tests/go`, or just `go`), one category
(`tests/go/json_marshal`) or a single file — and you may pass several, across
languages. Prefer a category: a suite is minutes, a category is seconds. On a
terminal `ok` is green and `FAILED` red (orange for a timeout); piped, there are
no escapes.

Read a saved log back with `testrunner summary <target>` — failures grouped by
`lang/category`, worst first, with a timeout column when a run had any. It takes
`php`, `tests/php`, `tests.php`, `php/bcmath` or a path, and with no exact match
merges every log with that prefix (so `summary php` covers the php categories
you saved), naming the files it merged.

`--help` lists the rest (`-j`, `--timeout`, `--results`, `--vybex`, `--verbose`).

`--json` writes `results/testrunner/run_<stamp>.json` and diffs it against the
previous run **of the same runtime**, naming regressions and newly-passing
tests; runs over different test sets compare only their overlap. It is opt-in —
a report per run buried `results/` in files nobody opened, and the diff needs
`--json` on the baseline run too.

`--save` writes the plain test log to `results/testrunner/saved/<target>.txt`
with `/` turned into `.` (`tests/php` → `tests.php.txt`), never coloured, for the
same stats scripts that read `<lang>.tests.txt`. **One file per target** —
`run tests/go tests/js --save` writes two logs, each with its own summary — and
**written as tests land**, so you can `tail -f` a long run. Works under
`--progress` too: the terminal shows the table, the file gets the cargo log.

## Running one file by hand

The point of extraction — no runner involved, full instrumentation:

```sh
./target/debug/vybex tests/go/json_marshal/marshal_bool_true.go
./target/debug/vybex -g          tests/go/json_marshal/marshal_bool_true.go  # step debugger
./target/debug/vybex --dump-ast  tests/go/json_marshal/marshal_bool_true.go
./target/debug/vybex -t          tests/go/json_marshal/marshal_bool_true.go  # bytecode trace
go run tests/go/json_marshal/marshal_bool_true.go                            # real Go
```

A test's verdict is its **exit code** — zero passes. That is the one mechanism
every target language and every foreign runtime shares (C, COBOL and Fortran
have no exceptions, but all can exit non-zero), and it is immune to the
`[vybex] Project …` banner that compilation prints on stdout.

## Warm mode

`vybex --worker` boots a VM once and runs one program per stdin line, resetting
between each. Roughly 90% of a `vybex <file>` invocation is VM setup — an empty
program costs 0.204s against 0.019s of process spawn — so paying it once per
worker instead of once per test is worth ~5× per test.

It is not a test feature: it is a warm execution host that happens to have the
runner as its first consumer.

```sh
printf 'a.go\nb.py\n' | ./target/debug/vybex --worker
##vybe-ready
...program output...
##vybe-result	ok
```

`--cold` exists so the two paths can be compared; they must agree on every
verdict, or the reset is leaking.

## Harnesses

`harness/<language>/check.<ext>` — real source in the language under test, the
way test262's `assert.js` is JavaScript. Edit `harness/go/check.go` with Go
tools; it is spliced into each emitted test at extraction time.

It prints its own `FAIL: want [...] got [...]` **before** failing. That is not
decoration: testecma relies on the thrown exception's message and 1,692 of its
2,158 failures come back as `RuntimeError: [object]`, saying nothing at all.

## Status

| language | extracted | passing under vybex |
|---|---|---|
| go | 6,393 | 5,634 (88.1%) |

Other languages need an emitter in `src/emit/` plus a `harness/<lang>/check.*`.
The `#[test] fn` suites (php, js, vb, pascal, python, java, cobol, ruby) also
need a second extractor shape — today only the `*_cases!` macro form is parsed.
