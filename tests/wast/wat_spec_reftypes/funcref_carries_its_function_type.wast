;; vybe-test: wast/wat_spec_reftypes/funcref_carries_its_function_type
;; vybe-test-mode: run
;;
;; KNOWN-FAILING. From `core/valid/types.rst` (function types) and the
;; `ref.test` / `ref.cast` rules in `core/exec/instructions.rst`.
;;
;; A function reference carries its FUNCTION TYPE. `ref.test (ref $t)`
;; answers 1 when the referenced function's type is `$t` (or a subtype), and
;; `ref.cast (ref $t)` succeeds on exactly those references and traps
;; otherwise. This holds wherever the reference came from — `ref.func`, a
;; table slot, an element segment, a global.
;;
;; Today no funcref carries a type from ANY source. `ref.func` builds an
;; object with `type_id: 0` and a single `__table_idx` property; the
;; element-segment path (`vm.rs`, the `Value::I32(func_idx)` arm) builds the
;; same shape. There is no functype on the value, so a concrete `(ref $t)`
;; test has nothing to match against and answers 0, and `ref.cast` traps with
;; "ref.cast failed: value is not thunk".
;;
;; The ABSTRACT tests already pass — `ref.test funcref` was fixed separately
;; (the reftype abbreviations resolved to a module type index). This file is
;; specifically about the CONCRETE type.

(module
  (type $thunk (func (result i32)))
  (type $taker (func (param i32) (result i32)))
  (type $void (func))

  (elem declare func $ten $inc $nothing)
  (func $ten (type $thunk) (i32.const 10))
  (func $inc (type $taker) (i32.add (local.get 0) (i32.const 1)))
  (func $nothing (type $void))

  (table $t 3 funcref)
  (elem (i32.const 0) $ten $inc $nothing)

  ;; ── ref.func ─────────────────────────────────────────────────────────
  (func (export "func_is_own_type") (result i32)
    (ref.test (ref $thunk) (ref.func $ten)))
  (func (export "func_is_other_type") (result i32)
    (ref.test (ref $taker) (ref.func $ten)))
  (func (export "taker_is_own_type") (result i32)
    (ref.test (ref $taker) (ref.func $inc)))
  (func (export "void_is_own_type") (result i32)
    (ref.test (ref $void) (ref.func $nothing)))
  ;; A funcref is still a funcref abstractly — this half already works, and is
  ;; here so a regression on it is visible alongside the concrete case.
  (func (export "func_is_funcref") (result i32)
    (ref.test funcref (ref.func $ten)))

  ;; ── through a TABLE ──────────────────────────────────────────────────
  (func (export "table_slot_is_own_type") (result i32)
    (ref.test (ref $thunk) (table.get $t (i32.const 0))))
  (func (export "table_slot_is_other_type") (result i32)
    (ref.test (ref $thunk) (table.get $t (i32.const 1))))

  ;; ── ref.cast succeeds on the right type ──────────────────────────────
  ;; Cast then call: if the cast carries the type through, the call works.
  (func (export "cast_then_call") (result i32)
    (call_ref $thunk (ref.cast (ref $thunk) (ref.func $ten))))
  (func (export "cast_wrong_type_traps") (result i32)
    (call_ref $taker (i32.const 1) (ref.cast (ref $taker) (ref.func $ten))))

  ;; ── call_ref itself, which needs the same type identity ──────────────
  (func (export "call_ref_direct") (result i32)
    (call_ref $thunk (ref.func $ten)))
  (func (export "call_ref_with_arg") (result i32)
    (call_ref $taker (i32.const 41) (ref.func $inc)))
)

;; ── a reference matches its OWN type and no other ─────────────────────
(assert_return (invoke "func_is_own_type") (i32.const 1))
(assert_return (invoke "func_is_other_type") (i32.const 0))
(assert_return (invoke "taker_is_own_type") (i32.const 1))
(assert_return (invoke "void_is_own_type") (i32.const 1))
(assert_return (invoke "func_is_funcref") (i32.const 1))

;; ── the same holds for a reference read out of a table ────────────────
(assert_return (invoke "table_slot_is_own_type") (i32.const 1))
(assert_return (invoke "table_slot_is_other_type") (i32.const 0))

;; ── ref.cast passes the type through to the call ──────────────────────
(assert_return (invoke "cast_then_call") (i32.const 10))
(assert_return (invoke "call_ref_direct") (i32.const 10))
(assert_return (invoke "call_ref_with_arg") (i32.const 42))
;; Casting to a type the reference does not have is a trap.
(assert_trap (invoke "cast_wrong_type_traps") "cast")
