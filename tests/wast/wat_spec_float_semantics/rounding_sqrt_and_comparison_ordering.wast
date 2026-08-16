;; vybe-test: wast/wat_spec_float_semantics/rounding_sqrt_and_comparison_ordering
;; vybe-test-mode: run
;;
;; From `op-fnearest` / `op-fceil` / `op-ffloor` / `op-ftrunc` / `op-fsqrt`
;; and the float comparison operators (core/exec/numerics.rst).
;;
;; Two rules carry everything here:
;;
;; 1. `nearest` rounds halfway cases to the EVEN integer, not away from zero.
;;    So nearest(1.5) = 2 but nearest(2.5) = 2 as well, and nearest(0.5) = 0.
;;    C's round() and JavaScript's Math.round() both disagree with this, which
;;    is where a wrong implementation usually comes from.
;;
;; 2. Every rounding operator preserves the sign of a zero RESULT, so anything
;;    in [-0.5, -0) rounds to NEGATIVE zero, and ceil of any value in (-1, 0)
;;    is -0 rather than +0. Signed zero is only observable through division,
;;    so those cases are read out with 1/z.
;;
;; The comparisons pin NaN's unorderedness: eq/lt/gt/le/ge are all FALSE when
;; either operand is NaN, and `ne` is the only one that is true — including
;; `ne(NaN, NaN)`. And +0 and -0 compare EQUAL despite differing bitwise.

(module
  (func (export "nearest") (param f64) (result f64) (f64.nearest (local.get 0)))
  (func (export "ceil") (param f64) (result f64) (f64.ceil (local.get 0)))
  (func (export "floor") (param f64) (result f64) (f64.floor (local.get 0)))
  (func (export "trunc") (param f64) (result f64) (f64.trunc (local.get 0)))
  (func (export "sqrt") (param f64) (result f64) (f64.sqrt (local.get 0)))
  (func (export "nearest32") (param f32) (result f32) (f32.nearest (local.get 0)))

  ;; Sign readers for zero results.
  (func (export "nearest_sign") (param f64) (result f64)
    (f64.div (f64.const 1) (f64.nearest (local.get 0))))
  (func (export "ceil_sign") (param f64) (result f64)
    (f64.div (f64.const 1) (f64.ceil (local.get 0))))
  (func (export "trunc_sign") (param f64) (result f64)
    (f64.div (f64.const 1) (f64.trunc (local.get 0))))
  (func (export "sqrt_sign") (param f64) (result f64)
    (f64.div (f64.const 1) (f64.sqrt (local.get 0))))

  (func (export "eq") (param f64 f64) (result i32) (f64.eq (local.get 0) (local.get 1)))
  (func (export "ne") (param f64 f64) (result i32) (f64.ne (local.get 0) (local.get 1)))
  (func (export "lt") (param f64 f64) (result i32) (f64.lt (local.get 0) (local.get 1)))
  (func (export "gt") (param f64 f64) (result i32) (f64.gt (local.get 0) (local.get 1)))
  (func (export "le") (param f64 f64) (result i32) (f64.le (local.get 0) (local.get 1)))
  (func (export "ge") (param f64 f64) (result i32) (f64.ge (local.get 0) (local.get 1)))

  (func (export "lt_s") (param i32 i32) (result i32) (i32.lt_s (local.get 0) (local.get 1)))
  (func (export "lt_u") (param i32 i32) (result i32) (i32.lt_u (local.get 0) (local.get 1)))
  (func (export "ge_s") (param i32 i32) (result i32) (i32.ge_s (local.get 0) (local.get 1)))
  (func (export "ge_u") (param i32 i32) (result i32) (i32.ge_u (local.get 0) (local.get 1)))
  (func (export "eqz") (param i32) (result i32) (i32.eqz (local.get 0)))
)

;; ── nearest: ties go to EVEN ──────────────────────────────────────────────
(assert_return (invoke "nearest" (f64.const 0.5)) (f64.const 0))
(assert_return (invoke "nearest" (f64.const 1.5)) (f64.const 2))
(assert_return (invoke "nearest" (f64.const 2.5)) (f64.const 2))
(assert_return (invoke "nearest" (f64.const 3.5)) (f64.const 4))
(assert_return (invoke "nearest" (f64.const 4.5)) (f64.const 4))
(assert_return (invoke "nearest" (f64.const -0.5)) (f64.const -0))
(assert_return (invoke "nearest" (f64.const -1.5)) (f64.const -2))
(assert_return (invoke "nearest" (f64.const -2.5)) (f64.const -2))
;; Non-ties round normally in both directions.
(assert_return (invoke "nearest" (f64.const 1.4)) (f64.const 1))
(assert_return (invoke "nearest" (f64.const 1.6)) (f64.const 2))
(assert_return (invoke "nearest" (f64.const -1.6)) (f64.const -2))
(assert_return (invoke "nearest32" (f32.const 2.5)) (f32.const 2))
(assert_return (invoke "nearest32" (f32.const 0.5)) (f32.const 0))

;; A tie that rounds to zero keeps the sign of the operand.
(assert_return (invoke "nearest_sign" (f64.const -0.5)) (f64.const -inf))
(assert_return (invoke "nearest_sign" (f64.const 0.5)) (f64.const inf))
(assert_return (invoke "nearest_sign" (f64.const -0.4)) (f64.const -inf))

;; ── ceil / floor / trunc, and their zero signs ───────────────────────────
(assert_return (invoke "ceil" (f64.const 1.1)) (f64.const 2))
(assert_return (invoke "ceil" (f64.const -1.1)) (f64.const -1))
(assert_return (invoke "floor" (f64.const 1.9)) (f64.const 1))
(assert_return (invoke "floor" (f64.const -1.1)) (f64.const -2))
(assert_return (invoke "trunc" (f64.const 1.9)) (f64.const 1))
(assert_return (invoke "trunc" (f64.const -1.9)) (f64.const -1))
;; ceil of anything in (-1, 0) is NEGATIVE zero, not +0.
(assert_return (invoke "ceil_sign" (f64.const -0.5)) (f64.const -inf))
(assert_return (invoke "ceil_sign" (f64.const -0.0)) (f64.const -inf))
(assert_return (invoke "trunc_sign" (f64.const -0.5)) (f64.const -inf))
(assert_return (invoke "trunc_sign" (f64.const 0.5)) (f64.const inf))
;; Infinities and NaN pass through every rounding operator unchanged.
(assert_return (invoke "ceil" (f64.const inf)) (f64.const inf))
(assert_return (invoke "floor" (f64.const -inf)) (f64.const -inf))
(assert_return (invoke "nearest" (f64.const inf)) (f64.const inf))
(assert_return (invoke "trunc" (f64.const nan)) (f64.const nan:canonical))

;; ── sqrt ─────────────────────────────────────────────────────────────────
(assert_return (invoke "sqrt" (f64.const 4)) (f64.const 2))
(assert_return (invoke "sqrt" (f64.const 0)) (f64.const 0))
(assert_return (invoke "sqrt" (f64.const inf)) (f64.const inf))
;; sqrt of a negative is NaN — but sqrt(-0) is -0, not NaN.
(assert_return (invoke "sqrt" (f64.const -1)) (f64.const nan:canonical))
(assert_return (invoke "sqrt" (f64.const -inf)) (f64.const nan:canonical))
(assert_return (invoke "sqrt" (f64.const -0)) (f64.const -0))
(assert_return (invoke "sqrt_sign" (f64.const -0)) (f64.const -inf))

;; ── NaN is unordered: five comparisons false, `ne` true ──────────────────
(assert_return (invoke "eq" (f64.const nan) (f64.const nan)) (i32.const 0))
(assert_return (invoke "ne" (f64.const nan) (f64.const nan)) (i32.const 1))
(assert_return (invoke "lt" (f64.const nan) (f64.const 1)) (i32.const 0))
(assert_return (invoke "gt" (f64.const nan) (f64.const 1)) (i32.const 0))
(assert_return (invoke "le" (f64.const nan) (f64.const 1)) (i32.const 0))
(assert_return (invoke "ge" (f64.const nan) (f64.const 1)) (i32.const 0))
(assert_return (invoke "lt" (f64.const 1) (f64.const nan)) (i32.const 0))
(assert_return (invoke "ge" (f64.const 1) (f64.const nan)) (i32.const 0))
;; le/ge are NOT "not gt"/"not lt" — NaN breaks that identity, which is the
;; whole reason both directions are listed here.
(assert_return (invoke "le" (f64.const nan) (f64.const nan)) (i32.const 0))
(assert_return (invoke "ge" (f64.const nan) (f64.const nan)) (i32.const 0))

;; +0 and -0 compare EQUAL despite differing bitwise.
(assert_return (invoke "eq" (f64.const 0) (f64.const -0)) (i32.const 1))
(assert_return (invoke "ne" (f64.const 0) (f64.const -0)) (i32.const 0))
(assert_return (invoke "lt" (f64.const -0) (f64.const 0)) (i32.const 0))
(assert_return (invoke "le" (f64.const -0) (f64.const 0)) (i32.const 1))
(assert_return (invoke "ge" (f64.const 0) (f64.const -0)) (i32.const 1))

;; Infinities are ordered against everything except NaN.
(assert_return (invoke "lt" (f64.const -inf) (f64.const inf)) (i32.const 1))
(assert_return (invoke "eq" (f64.const inf) (f64.const inf)) (i32.const 1))
(assert_return (invoke "lt" (f64.const -inf) (f64.const -1e308)) (i32.const 1))

;; ── The same bits, read signed and unsigned, order differently ───────────
(assert_return (invoke "lt_s" (i32.const -1) (i32.const 1)) (i32.const 1))
(assert_return (invoke "lt_u" (i32.const -1) (i32.const 1)) (i32.const 0))
(assert_return (invoke "lt_s" (i32.const -2147483648) (i32.const 0)) (i32.const 1))
(assert_return (invoke "lt_u" (i32.const -2147483648) (i32.const 0)) (i32.const 0))
(assert_return (invoke "ge_s" (i32.const -1) (i32.const 0)) (i32.const 0))
(assert_return (invoke "ge_u" (i32.const -1) (i32.const 0)) (i32.const 1))
(assert_return (invoke "lt_s" (i32.const -2) (i32.const -1)) (i32.const 1))
(assert_return (invoke "lt_u" (i32.const -2) (i32.const -1)) (i32.const 1))
(assert_return (invoke "eqz" (i32.const 0)) (i32.const 1))
(assert_return (invoke "eqz" (i32.const -2147483648)) (i32.const 0))
