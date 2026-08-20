;; traps.wat — the WASM trap, and how it differs from an exception
;; Run: cargo run --bin vybex -- examples/wast/traps.wat
;;
;; A TRAP is not an exception. An exception is a value: it carries a tag and a
;; payload, it searches for a `try_table` handler, and a program can catch it,
;; inspect it or swallow it. A trap is the machine refusing to continue — no
;; value, no tag, nothing to catch. Per the WASM spec a trap is OUTSIDE the
;; exception system, so no handler in the module can intercept one.
;;
;; That is exactly why a test harness wants it: a trap cannot be swallowed. See
;; `exceptions.wat` for the value-carrying side of the story.
;;
;; This program logs its guarded checks, then traps ON PURPOSE. A NON-ZERO exit
;; is the demonstration, not a failure of the example.
;;
;;   ── KNOWN DIVERGENCE FROM THE SPEC, measured 2026-08-05 ──
;;
;; Vybe does not yet emit `Op::UNREACHABLE`; no front end can produce it, though
;; the VM defines and executes it. `languages/wast/src/walker.rs` lowers
;; `unreachable` to `StmtKind::Throw { expr: None }` instead. Two consequences:
;;
;;   1. STATEMENT position (below) ends the program with a non-zero exit, so it
;;      LOOKS right — but it reports `RuntimeError: null` rather than
;;      `trap: unreachable executed`, and being a throw it is CATCHABLE.
;;      Measured: `(block $done (try_table (catch_all $done) unreachable))`
;;      swallows it and the program carries on and exits 0. No spec-conformant
;;      engine can do that.
;;
;;   2. EXPRESSION position — a folded `(unreachable)` used as a value — is
;;      compiled to `Expression::null()`, i.e. NOTHING. `$never_ok` below
;;      returns null and execution carries on. A WAT program whose failure path
;;      is written that way exits 0 and reports success.
;;
;; `$never_ok` is exported but deliberately NOT called: calling it today proves
;; nothing, because it does not trap. Call it once the lowering is fixed and it
;; must end the program.

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  ;; A guarded operation: trap rather than return a wrong answer. This is the
  ;; idiom `array.get` and `i32.div_s` use internally — the check is the
  ;; contract, and violating it is not a recoverable condition.
  (func $checked_div (export "checked_div") (param $a i32) (param $b i32) (result i32)
    local.get $b
    i32.eqz
    if
      unreachable
    end
    local.get $a
    local.get $b
    i32.div_s)

  ;; Same shape for a bounds check.
  (func $checked_index (export "checked_index") (param $i i32) (param $len i32) (result i32)
    local.get $i
    i32.const 0
    i32.lt_s
    if
      unreachable
    end
    local.get $i
    local.get $len
    i32.ge_s
    if
      unreachable
    end
    local.get $i)

  ;; DIVERGENCE #2 — a folded `(unreachable)` in VALUE position. Today this
  ;; returns null and the caller keeps running. Not called by `$demo`.
  (func $never_ok (export "never_ok") (result i32)
    (unreachable))

  (func $demo (export "demo")
    ;; Every guard holds — nothing traps, ordinary values come back.
    (call $log (call $checked_div (i32.const 84) (i32.const 2)))     ;; 42
    (call $log (call $checked_index (i32.const 3) (i32.const 10)))   ;; 3
    (call $log (call $checked_index (i32.const 9) (i32.const 10)))   ;; 9

    ;; Now violate one. Nothing after this line runs, and no handler anywhere
    ;; in the module can stop it.
    (call $log (call $checked_div (i32.const 1) (i32.const 0)))
    (call $log (i32.const 1234)))                                    ;; never reached

  (start $demo)
)
