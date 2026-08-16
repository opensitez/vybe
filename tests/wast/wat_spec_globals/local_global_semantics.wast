;; vybe-test: wast/wat_spec_globals/local_global_semantics
;; vybe-test-mode: run
;;
;; From `core/syntax/modules.rst` (globals) and `core/exec/modules.rst`
;; (instantiation).
;;
;; This file is the GUARD on globals that already work. It exists separately
;; from the imported-global and validation files on purpose: the cheap way to
;; make a validation suite pass is to start rejecting things that are legal,
;; and the cheap way to make a linking fix look good is to disturb the local
;; path. This file is what makes either regression visible immediately.
;;
;; Covered: mutable and immutable globals, a reference-typed global, `start`
;; running at instantiation before any exported call, mutation visible to
;; later reads including across a call boundary, and the constant-expression
;; forms that are always legal (`t.const`, `ref.null`, `ref.func`).

(module
  (global $mut_i32 (mut i32) (i32.const 100))
  (global $imm_i32 i32 (i32.const 5))
  (global $mut_i64 (mut i64) (i64.const -1))
  (global $mut_f64 (mut f64) (f64.const 1.5))
  (global $nullref (mut funcref) (ref.null func))
  (global $extnull (mut externref) (ref.null extern))
  (elem declare func $ten)
  (func $ten (result i32) (i32.const 10))
  ;; A `ref.func` initialiser. This once read back null, but that turned out
  ;; to be state left over from a previously compiled module rather than a
  ;; defect in the initialiser — it is guarded here so the distinction stays
  ;; visible if it ever regresses for real.
  (global $funcref_init funcref (ref.func $ten))
  (table 1 funcref)
  (type $thunk (func (result i32)))

  ;; `start` mutates a global, so its effect is observable afterwards.
  (global $started (mut i32) (i32.const 0))
  (func $start (global.set $started (i32.const 1)))
  (start $start)

  (func (export "mut_i32") (result i32) (global.get $mut_i32))
  (func (export "imm_i32") (result i32) (global.get $imm_i32))
  (func (export "mut_i64") (result i64) (global.get $mut_i64))
  (func (export "mut_f64") (result f64) (global.get $mut_f64))
  (func (export "started") (result i32) (global.get $started))
  (func (export "set_i32") (param i32) (global.set $mut_i32 (local.get 0)))
  (func (export "set_i64") (param i64) (global.set $mut_i64 (local.get 0)))
  (func (export "set_f64") (param f64) (global.set $mut_f64 (local.get 0)))
  (func (export "nullref_is_null") (result i32) (ref.is_null (global.get $nullref)))
  (func (export "extnull_is_null") (result i32) (ref.is_null (global.get $extnull)))
  (func (export "funcref_is_null") (result i32) (ref.is_null (global.get $funcref_init)))
  ;; A funcref held in a global is genuinely callable.
  (func (export "call_from_global") (result i32)
    (table.set (i32.const 0) (global.get $funcref_init))
    (call_indirect (type $thunk) (i32.const 0)))
  ;; A read inside a CALLED function sees the current value, not one captured
  ;; at instantiation.
  (func $reader (result i32) (global.get $mut_i32))
  (func (export "read_via_call") (result i32) (call $reader))
  ;; A global survives being written from inside a loop.
  (func (export "accumulate") (param i32) (result i32)
    (global.set $mut_i32 (i32.const 0))
    (block $done
      (loop $again
        (br_if $done (i32.eqz (local.get 0)))
        (global.set $mut_i32 (i32.add (global.get $mut_i32) (local.get 0)))
        (local.set 0 (i32.sub (local.get 0) (i32.const 1)))
        (br $again)))
    (global.get $mut_i32))
)

;; ── Initialisers ran ───────────────────────────────────────────────────
(assert_return (invoke "mut_i32") (i32.const 100))
(assert_return (invoke "imm_i32") (i32.const 5))
(assert_return (invoke "mut_i64") (i64.const -1))
(assert_return (invoke "mut_f64") (f64.const 1.5))
(assert_return (invoke "nullref_is_null") (i32.const 1))
(assert_return (invoke "extnull_is_null") (i32.const 1))
(assert_return (invoke "funcref_is_null") (i32.const 0))
(assert_return (invoke "call_from_global") (i32.const 10))

;; ── `start` ran at instantiation, before any exported call ─────────────
(assert_return (invoke "started") (i32.const 1))

;; ── Mutation is visible to later reads, and through a call ─────────────
(invoke "set_i32" (i32.const 42))
(assert_return (invoke "mut_i32") (i32.const 42))
(assert_return (invoke "read_via_call") (i32.const 42))
(invoke "set_i32" (i32.const -2147483648))
(assert_return (invoke "mut_i32") (i32.const -2147483648))
(assert_return (invoke "read_via_call") (i32.const -2147483648))
(invoke "set_i64" (i64.const 9223372036854775807))
(assert_return (invoke "mut_i64") (i64.const 9223372036854775807))
(invoke "set_f64" (f64.const -0))
(assert_return (invoke "mut_f64") (f64.const -0))

;; ── A global written in a loop keeps its value across iterations ───────
(assert_return (invoke "accumulate" (i32.const 0)) (i32.const 0))
(assert_return (invoke "accumulate" (i32.const 4)) (i32.const 10))
(assert_return (invoke "accumulate" (i32.const 100)) (i32.const 5050))
;; ...and the last write is still there afterwards.
(assert_return (invoke "mut_i32") (i32.const 5050))

;; ── The immutable global is unchanged by all of the above ──────────────
(assert_return (invoke "imm_i32") (i32.const 5))
