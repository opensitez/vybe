;; vybe-test: wast/wat_call_indirect/multi_result_callee_lands_as_separate_values
;; vybe-test-mode: run
;;
;; WASM multi-value: a call leaves N values on the stack, and the walker models
;; that by DESTRUCTURING the callee's packed result into N fresh temps and
;; pushing them individually (`land_instr_value`). A direct `call` to a
;; multi-result function did that; `call_indirect` pushed the packed tuple as
;; ONE stack entry instead.
;;
;; Nothing downstream noticed while the values were only ever consumed by an
;; arithmetic instruction — `(i32.add (call_indirect …))` works either way,
;; because the add pops what it finds. It shows up when the function RETURNS
;; them: `apply_multi_value_return` looks for N contiguous value-statements at
;; the tail, a single packed push is one statement, so it declined and the
;; function returned nothing at all. The assertion then failed with no hint
;; that the call had been fine and the RETURN was the problem.
;;
;; Both spellings of the type use are exercised, because the result count
;; reaches the same lowering from `(type $t)` and from an inline `(result …)`.

(module
  (type $r2 (func (result i32 i32)))
  (type $r3 (func (result i32 i64 f64)))
  (type $p2r2 (func (param i32 i32) (result i32 i32)))

  (table 4 funcref)
  (elem (i32.const 0) $two $three $swap $one)

  (func $two (result i32 i32) (i32.const 11) (i32.const 22))
  (func $three (result i32 i64 f64)
    (i32.const 1) (i64.const 2) (f64.const 3.5))
  (func $swap (param $a i32) (param $b i32) (result i32 i32)
    (local.get $b) (local.get $a))
  (func $one (result i32) (i32.const 99))

  ;; ── the case that was broken: return them straight out ─────────────
  (func (export "named_pair") (param $i i32) (result i32 i32)
    (call_indirect (type $r2) (local.get $i)))
  (func (export "inline_pair") (param $i i32) (result i32 i32)
    (call_indirect (result i32 i32) (local.get $i)))

  ;; Three results, mixed types — the temps must keep their order and their
  ;; values, not just their count.
  (func (export "named_triple") (param $i i32) (result i32 i64 f64)
    (call_indirect (type $r3) (local.get $i)))

  ;; Params and results together: `$swap` returns its arguments reversed, so a
  ;; lost or reordered temp is visible in the answer rather than in the count.
  (func (export "swapped") (param $i i32) (result i32 i32)
    (call_indirect (type $p2r2) (i32.const 5) (i32.const 6) (local.get $i)))

  ;; ── the controls ───────────────────────────────────────────────────
  ;; A DIRECT call in the same position always worked, and must keep working.
  (func (export "direct_pair") (result i32 i32) (call $two))

  ;; Consumed rather than returned — this passed before the fix too, and is
  ;; here so a change that broke the consuming path would be caught.
  (func (export "summed") (param $i i32) (result i32)
    (i32.add (call_indirect (type $r2) (local.get $i))))

  ;; A single-result call_indirect must still land as ONE value.
  (func (export "single") (param $i i32) (result i32)
    (call_indirect (result i32) (local.get $i)))

  ;; …and the pair must still be usable after other work happens between the
  ;; call and the return, which is what the temps are for.
  (func (export "pair_then_work") (param $i i32) (result i32 i32)
    (local $scratch i32)
    (call_indirect (type $r2) (local.get $i))
    (local.set $scratch (i32.const 7))
    ;; the two results are still the top of the stack
  )
)

;; Slot 0 `$two` (0→2), slot 1 `$three` (0→3), slot 2 `$swap` (2→2),
;; slot 3 `$one` (0→1).
(assert_return (invoke "named_pair" (i32.const 0)) (i32.const 11) (i32.const 22))
(assert_return (invoke "inline_pair" (i32.const 0)) (i32.const 11) (i32.const 22))
(assert_return (invoke "named_triple" (i32.const 1))
               (i32.const 1) (i64.const 2) (f64.const 3.5))
(assert_return (invoke "swapped" (i32.const 2)) (i32.const 6) (i32.const 5))
(assert_return (invoke "pair_then_work" (i32.const 0)) (i32.const 11) (i32.const 22))

(assert_return (invoke "direct_pair") (i32.const 11) (i32.const 22))
(assert_return (invoke "summed" (i32.const 0)) (i32.const 33))
(assert_return (invoke "single" (i32.const 3)) (i32.const 99))
