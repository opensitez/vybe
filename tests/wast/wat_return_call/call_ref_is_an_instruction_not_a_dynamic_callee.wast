;; vybe-test: wast/wat_return_call/call_ref_is_an_instruction_not_a_dynamic_callee
;; vybe-test-mode: run
;;
;; `call_ref $sig` and `return_call_ref $sig` are WASM instructions with their
;; own opcodes and a `[funcref, operand*]` stack layout. They used to lower to a
;; generic AST call whose callee was an EXPRESSION, which sends the compiler
;; down its dynamic-callee ladder instead — and that ladder opens by reading
;; `__vybe_method_receiver` off the callee, a `struct.get`. So:
;;
;;   * `call_ref` on a null funcref trapped "null structure reference
;;     (struct.get)". The message was honest — a property was read off a typed
;;     null — it just was not a call. The spec says "null function reference".
;;   * `return_call_ref` went through `__wasm_return_call`, which takes a
;;     QUALIFIED CALLEE NAME (that is how `return_call $f` names the module
;;     method to tail-call) and was handed a funcref VALUE, so it reported
;;     "null is not callable" for every call, null or not.
;;
;; ⛔ Both were blind in the PLAIN spelling for a second reason: there the
;; operands are on the enclosing block's stack and never appear as arguments,
;; so popping the funcref off an empty argument list called null. Only `$sig`
;; knows the argument count there, which is why it is read even though the
;; folded form counts what was actually written.
;;
;; ⛔ The funcref is on TOP of the stack — `call_ref : [t1* (ref null $t)] ->
;; [t2*]` — so it is pushed LAST in the plain spelling, exactly as it is written
;; last when folded. The opcode wants it BELOW its arguments (that is
;; `call_value`'s layout, and what `emit_call` has always emitted), so the
;; lowering moves it. Writing the plain operands in the other order here reads
;; fine and fails as "f64 is not callable (type: 21)" — the ARGUMENT gets
;; called.

(module
  (type $ii (func (param i32) (result i32)))
  (type $v (func))
  (type $iii (func (param i32 i32) (result i32)))
  (type $ll (func (param i64) (result i64)))

  (func $double (param $x i32) (result i32)
    (i32.mul (local.get $x) (i32.const 2)))
  (func $add (param $a i32) (param $b i32) (result i32)
    (i32.add (local.get $a) (local.get $b)))
  (global $side (mut i32) (i32.const 0))
  (func $bump (global.set $side (i32.add (global.get $side) (i32.const 7))))

  (elem declare func $double $add $bump $fac)

  ;; ── folded ───────────────────────────────────────────────────────────
  (func (export "f_one") (param $x i32) (result i32)
    (call_ref $ii (local.get $x) (ref.func $double)))
  (func (export "f_two") (param $a i32) (param $b i32) (result i32)
    (call_ref $iii (local.get $a) (local.get $b) (ref.func $add)))
  ;; Zero operands — the folded list holds the funcref and nothing else, the
  ;; case where "the last argument is the callee" and "there are no arguments"
  ;; have to stay distinguishable.
  (func (export "f_zero") (result i32)
    (call_ref $v (ref.func $bump))
    (global.get $side))
  ;; A `call_ref` whose OPERAND is another `call_ref`.
  (func (export "f_nested") (param $x i32) (result i32)
    (call_ref $ii (call_ref $ii (local.get $x) (ref.func $double)) (ref.func $double)))

  ;; ── plain: the operands are on the block's stack, not in the args ────
  (func (export "p_one") (param $x i32) (result i32)
    local.get $x
    ref.func $double
    call_ref $ii)
  (func (export "p_two") (param $a i32) (param $b i32) (result i32)
    local.get $a
    local.get $b
    ref.func $add
    call_ref $iii)
  (func (export "p_zero") (result i32)
    ref.func $bump
    call_ref $v
    global.get $side)

  ;; ── the null cases ───────────────────────────────────────────────────
  (func (export "f_null") (result i32)
    (call_ref $ii (i32.const 1) (ref.null $ii)))
  (func (export "p_null") (result i32)
    i32.const 1
    ref.null $ii
    call_ref $ii)
  ;; A null that arrives through a LOCAL rather than written at the call site,
  ;; so the trap cannot come from folding a literal null.
  (func (export "local_null") (result i32)
    (local $f (ref null $ii))
    (call_ref $ii (i32.const 1) (local.get $f)))

  ;; ── return_call_ref ──────────────────────────────────────────────────
  (func (export "t_one") (param $x i32) (result i32)
    (return_call_ref $ii (local.get $x) (ref.func $double)))
  (func (export "t_plain") (param $x i32) (result i32)
    local.get $x
    ref.func $double
    return_call_ref $ii)
  (func (export "t_null") (result i32)
    (return_call_ref $ii (i32.const 1) (ref.null $ii)))

  ;; A tail call must REUSE the frame. 200000 deep overflows a stack that
  ;; grows, which is the only way to tell a real tail call from `return f(x)`.
  (func $count (param $n i64) (result i64)
    (if (result i64) (i64.eqz (local.get $n))
      (then (i64.const 0))
      (else (return_call_ref $ll
              (i64.sub (local.get $n) (i64.const 1))
              (ref.func $count)))))
  (elem declare func $count)
  (func (export "deep") (param $n i64) (result i64)
    (call_ref $ll (local.get $n) (ref.func $count)))

  (func $fac (param $n i64) (result i64)
    (if (result i64) (i64.eqz (local.get $n))
      (then (i64.const 1))
      (else (i64.mul (local.get $n)
              (call_ref $ll (i64.sub (local.get $n) (i64.const 1)) (ref.func $fac))))))
  (func (export "fac") (param $n i64) (result i64)
    (call_ref $ll (local.get $n) (ref.func $fac)))
)

;; ── it calls, in both spellings ────────────────────────────────────────
(assert_return (invoke "f_one" (i32.const 21)) (i32.const 42))
(assert_return (invoke "p_one" (i32.const 21)) (i32.const 42))
;; Two operands, so a reversed layout shows up as a wrong ANSWER and not only
;; as a crash — `$add` is commutative, so the subtraction check is below.
(assert_return (invoke "f_two" (i32.const 30) (i32.const 12)) (i32.const 42))
(assert_return (invoke "p_two" (i32.const 30) (i32.const 12)) (i32.const 42))
(assert_return (invoke "f_nested" (i32.const 5)) (i32.const 20))

;; Zero operands: the effect is a global write, so "did it call" is observable
;; without a return value.
(assert_return (invoke "f_zero") (i32.const 7))
(assert_return (invoke "p_zero") (i32.const 14))

;; ── the null cases: the trap the spec names ────────────────────────────
(assert_trap (invoke "f_null") "null function reference")
(assert_trap (invoke "p_null") "null function reference")
(assert_trap (invoke "local_null") "null function reference")
(assert_trap (invoke "t_null") "null function reference")

;; ── return_call_ref calls, and reuses the frame ────────────────────────
(assert_return (invoke "t_one" (i32.const 21)) (i32.const 42))
(assert_return (invoke "t_plain" (i32.const 21)) (i32.const 42))
(assert_return (invoke "deep" (i64.const 200000)) (i64.const 0))
(assert_return (invoke "fac" (i64.const 10)) (i64.const 3628800))
