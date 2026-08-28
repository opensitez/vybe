;; vybe-test: wast/wat_call_indirect/inline_type_use_states_the_shape
;; vybe-test-mode: run
;;
;; A type use is EITHER `(type $sig)` naming a declared function type OR the
;; inline `(param …)* (result …)*` spelling of the same shape. Both are legal
;; everywhere a type use appears, and the inline form needs no `(type …)` at
;; all — `call_indirect (result i32)` is how most of the spec suite writes it.
;;
;; Only the named form was read. The inline one came out as 0 params and 0
;; results, so the VM compared the funcref's real shape against 0→0 and
;; rejected a perfectly good callee: "indirect call type mismatch (callee 0→1,
;; expected 0→0)". The trap named the callee correctly and the expectation
;; wrongly, which reads like a bad table entry rather than a lost immediate.
;;
;; The declared-type forms below are the control: they were always right, and
;; must stay right — the shape is now read by one function for both spellings,
;; so a mistake in it would take them down too.

(module
  (type $ret_i32 (func (result i32)))
  (type $two_args (func (param i32 i32) (result i32)))

  (table 8 funcref)
  (elem (i32.const 0) $const7 $add $swallow $two_results)

  (func $const7 (result i32) (i32.const 7))
  (func $add (param $a i32) (param $b i32) (result i32)
    (i32.add (local.get $a) (local.get $b)))
  (func $swallow (param $a i32))
  (func $two_results (result i32 i32) (i32.const 11) (i32.const 22))

  ;; ── inline type use ────────────────────────────────────────────────
  ;; No `(type …)`: the shape is stated by `(result i32)` alone.
  (func (export "inline_result") (param $i i32) (result i32)
    (call_indirect (result i32) (local.get $i)))

  ;; Params and results together, inline.
  (func (export "inline_params_result") (param $i i32) (result i32)
    (call_indirect (param i32) (param i32) (result i32)
      (i32.const 30) (i32.const 12) (local.get $i)))

  ;; Several params in ONE `(param …)` — the count is the number of value
  ;; types, not the number of clauses.
  (func (export "inline_multi_param") (param $i i32) (result i32)
    (call_indirect (param i32 i32) (result i32)
      (i32.const 40) (i32.const 2) (local.get $i)))

  ;; A named param binds nothing here but is legal and must not be counted as
  ;; a value type.
  (func (export "inline_named_param") (param $i i32) (result i32)
    (call_indirect (param $x i32) (param $y i32) (result i32)
      (i32.const 5) (i32.const 6) (local.get $i)))

  ;; No results at all — the inline form of a 1→0 callee.
  (func (export "inline_no_result") (param $i i32)
    (call_indirect (param i32) (i32.const 99) (local.get $i)))

  ;; Multi-value: two results stated inline. Consumed by an `i32.add` rather
  ;; than returned as a pair, because RETURNING a call_indirect's results
  ;; straight out of a `(result i32 i32)` function is separately broken —
  ;; `apply_multi_value_return` wants N contiguous value-statements at the tail
  ;; and a single call that pushes N is one statement, so the tuple return is
  ;; never applied. That is true of the `(type $t)` spelling too, so it is not
  ;; this fix's business; asserting it here would report a multi-value-return
  ;; bug as a type-use bug. The add still proves the immediate is read as TWO
  ;; results: at 0 the VM would reject the 0→2 callee outright.
  (func (export "inline_two_results_sum") (param $i i32) (result i32)
    (i32.add (call_indirect (result i32 i32) (local.get $i))))

  ;; ── the control: the declared-type spelling ────────────────────────
  (func (export "named_result") (param $i i32) (result i32)
    (call_indirect (type $ret_i32) (local.get $i)))
  (func (export "named_params_result") (param $i i32) (result i32)
    (call_indirect (type $two_args) (i32.const 1) (i32.const 2) (local.get $i)))
)

;; Slot 0 is `$const7` (0→1), slot 1 `$add` (2→1), slot 2 `$swallow` (1→0),
;; slot 3 `$two_results` (0→2).
(assert_return (invoke "inline_result" (i32.const 0)) (i32.const 7))
(assert_return (invoke "inline_params_result" (i32.const 1)) (i32.const 42))
(assert_return (invoke "inline_multi_param" (i32.const 1)) (i32.const 42))
(assert_return (invoke "inline_named_param" (i32.const 1)) (i32.const 11))
(invoke "inline_no_result" (i32.const 2))
(assert_return (invoke "inline_two_results_sum" (i32.const 3)) (i32.const 33))

(assert_return (invoke "named_result" (i32.const 0)) (i32.const 7))
(assert_return (invoke "named_params_result" (i32.const 1)) (i32.const 3))

;; A shape that genuinely disagrees still traps — reading the inline form must
;; not amount to accepting anything. Slot 1 is `$add`, which is 2→1, so a 0→1
;; call through it is a real mismatch.
(assert_trap (invoke "inline_result" (i32.const 1)) "indirect call type mismatch")
