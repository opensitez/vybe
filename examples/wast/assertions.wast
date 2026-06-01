;; assertions.wast — WAST script commands: assert_return, assert_trap,
;; assert_invalid, assert_malformed, assert_unlinkable, register, invoke, get
;; Run: cargo run -p vybex -- examples/wast/assertions.wast

;; ── Module under test ────────────────────────────────────────────────────────

(module $math
  (func (export "add")  (param i32 i32) (result i32) local.get 0 local.get 1 i32.add)
  (func (export "sub")  (param i32 i32) (result i32) local.get 0 local.get 1 i32.sub)
  (func (export "mul")  (param i32 i32) (result i32) local.get 0 local.get 1 i32.mul)
  (func (export "div")  (param i32 i32) (result i32) local.get 0 local.get 1 i32.div_s)
  (func (export "neg")  (param i32)     (result i32) i32.const 0 local.get 0 i32.sub)
  (func (export "boom") unreachable)
  (func (export "div0") (param i32)     (result i32) local.get 0 i32.const 0 i32.div_s)
  (global (export "answer") i32 (i32.const 42))
)

;; ── assert_return ─────────────────────────────────────────────────────────────

(assert_return (invoke $math "add" (i32.const 1)  (i32.const 2))  (i32.const 3))
(assert_return (invoke $math "add" (i32.const 0)  (i32.const 0))  (i32.const 0))
(assert_return (invoke $math "add" (i32.const -1) (i32.const 1))  (i32.const 0))
(assert_return (invoke $math "sub" (i32.const 10) (i32.const 3))  (i32.const 7))
(assert_return (invoke $math "mul" (i32.const 6)  (i32.const 7))  (i32.const 42))
(assert_return (invoke $math "div" (i32.const 10) (i32.const 2))  (i32.const 5))
(assert_return (invoke $math "neg" (i32.const 5))                 (i32.const -5))
(assert_return (invoke $math "neg" (i32.const 0))                 (i32.const 0))

;; ── assert_return with floats ─────────────────────────────────────────────────

(module $floats
  (func (export "pi")    (result f64) f64.const 3.141592653589793)
  (func (export "sqrt2") (result f64) f64.const 1.4142135623730951)
  (func (export "inf")   (result f64) f64.const inf)
  (func (export "nan")   (result f64) f64.const nan)
  (func (export "fadd")  (param f64 f64) (result f64) local.get 0 local.get 1 f64.add)
  (func (export "fmul")  (param f32 f32) (result f32) local.get 0 local.get 1 f32.mul)
)

(assert_return (invoke $floats "pi")                                    (f64.const 3.141592653589793))
(assert_return (invoke $floats "inf")                                   (f64.const inf))
(assert_return (invoke $floats "nan")                                   (f64.const nan:canonical))
(assert_return (invoke $floats "fadd" (f64.const 1.5) (f64.const 2.5)) (f64.const 4.0))
(assert_return (invoke $floats "fmul" (f32.const 2.0) (f32.const 3.0)) (f32.const 6.0))

;; ── assert_return with multiple results ──────────────────────────────────────

(module $multi
  (func (export "swap") (param i32 i32) (result i32 i32) local.get 1 local.get 0)
  (func (export "dup")  (param i32)     (result i32 i32) local.get 0 local.get 0)
)

(assert_return (invoke $multi "swap" (i32.const 1) (i32.const 2)) (i32.const 2) (i32.const 1))
(assert_return (invoke $multi "dup"  (i32.const 7))               (i32.const 7) (i32.const 7))

;; ── assert_trap ───────────────────────────────────────────────────────────────

(assert_trap (invoke $math "boom")                    "unreachable")
(assert_trap (invoke $math "div0" (i32.const 1))      "integer divide by zero")

;; ── assert_return with get (global) ──────────────────────────────────────────

(assert_return (get $math "answer") (i32.const 42))

;; ── register + cross-module import ───────────────────────────────────────────

(register "math" $math)

(module $consumer
  (import "math" "add" (func $add (param i32 i32) (result i32)))
  (import "math" "mul" (func $mul (param i32 i32) (result i32)))
  (func (export "dot") (param $a i32) (param $b i32) (param $c i32) (param $d i32) (result i32)
    ;; dot product of (a,b)·(c,d) = a*c + b*d
    (i32.add
      (call $mul (local.get $a) (local.get $c))
      (call $mul (local.get $b) (local.get $d))))
)

(assert_return
  (invoke $consumer "dot" (i32.const 1) (i32.const 2) (i32.const 3) (i32.const 4))
  (i32.const 11))  ;; 1*3 + 2*4 = 11

;; ── assert_invalid ────────────────────────────────────────────────────────────

;; Type mismatch: function says result i32 but leaves f32 on stack
(assert_invalid
  (module (func (result i32) f32.const 1.0))
  "type mismatch")

;; Stack underflow: result expected but nothing pushed
(assert_invalid
  (module (func (result i32)))
  "type mismatch")

;; Unknown local index
(assert_invalid
  (module (func (result i32) local.get 5))
  "unknown local")

;; ── assert_malformed ──────────────────────────────────────────────────────────

;; Binary module with wrong magic
(assert_malformed
  (module binary "\00\61\73\6d")
  "unexpected end")

;; Quoted module with syntax error
(assert_malformed
  (module quote "(module (func (result i32) i32.const))")
  "unexpected token")

;; ── assert_unlinkable ─────────────────────────────────────────────────────────

;; Import that doesn't exist
(assert_unlinkable
  (module (import "nonexistent" "fn" (func)))
  "unknown import")

;; ── assert_exhaustion ─────────────────────────────────────────────────────────

(module $inf_loop
  (func $recurse (export "recurse") call $recurse)
)

(assert_exhaustion (invoke $inf_loop "recurse") "call stack exhausted")
