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

`--help` lists the rest (`-j`, `--timeout`, `--results`, `--vybex`, `--verbose`).

Reports land in `results/testrunner/run_<stamp>.json`, and each run is diffed
against the previous run **of the same runtime**, naming regressions and
newly-passing tests. Runs over different test sets compare only their overlap.

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
