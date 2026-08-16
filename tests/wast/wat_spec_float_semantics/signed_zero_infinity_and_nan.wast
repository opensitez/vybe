;; vybe-test: wast/wat_spec_float_semantics/signed_zero_infinity_and_nan
;; vybe-test-mode: run
;;
;; From the spec's `op-fadd` / `op-fmin` / `op-fmax` case lists
;; (core/exec/numerics.rst) and the NaN-propagation rule at `aux-nans`.
;;
;; Each of these operators is defined as an ORDERED list of cases, and almost
;; every case exists to pin down a sign that IEEE arithmetic alone would leave
;; ambiguous. They are the cases a naive implementation gets wrong while
;; producing numerically correct answers everywhere else:
;;
;;   fadd(±0, ∓0) = +0      -- but fadd(-0, -0) = -0
;;   fadd(±q, ∓q) = +0      -- cancellation is POSITIVE zero
;;   fmin(±0, ∓0) = -0      -- fmin is not "whichever compares smaller":
;;                             -0 and +0 compare EQUAL, so the rule is explicit
;;   fmax(±0, ∓0) = +0
;;   fmin(-∞, x)  = -∞      -- checked before the NaN-free comparison
;;   fmax(NaN, x) = NaN     -- NaN wins over infinity; min/max do NOT
;;                             propagate the non-NaN operand the way the
;;                             IEEE-754 minNum/maxNum recommendation does
;;
;; Signed zero is observable only through division, so every zero-sign claim
;; below is read out with 1/z rather than compared directly — `assert_return`
;; on a zero would pass for either sign.

(module
  (func (export "add") (param f64 f64) (result f64) (f64.add (local.get 0) (local.get 1)))
  (func (export "sub") (param f64 f64) (result f64) (f64.sub (local.get 0) (local.get 1)))
  (func (export "mul") (param f64 f64) (result f64) (f64.mul (local.get 0) (local.get 1)))
  (func (export "div") (param f64 f64) (result f64) (f64.div (local.get 0) (local.get 1)))
  (func (export "min") (param f64 f64) (result f64) (f64.min (local.get 0) (local.get 1)))
  (func (export "max") (param f64 f64) (result f64) (f64.max (local.get 0) (local.get 1)))
  (func (export "min32") (param f32 f32) (result f32) (f32.min (local.get 0) (local.get 1)))
  (func (export "max32") (param f32 f32) (result f32) (f32.max (local.get 0) (local.get 1)))
  ;; Reads the SIGN of a zero result: 1/+0 = +inf, 1/-0 = -inf.
  (func (export "add_sign") (param f64 f64) (result f64)
    (f64.div (f64.const 1) (f64.add (local.get 0) (local.get 1))))
  (func (export "min_sign") (param f64 f64) (result f64)
    (f64.div (f64.const 1) (f64.min (local.get 0) (local.get 1))))
  (func (export "max_sign") (param f64 f64) (result f64)
    (f64.div (f64.const 1) (f64.max (local.get 0) (local.get 1))))
  (func (export "copysign") (param f64 f64) (result f64)
    (f64.copysign (local.get 0) (local.get 1)))
  (func (export "neg") (param f64) (result f64) (f64.neg (local.get 0)))
  (func (export "abs") (param f64) (result f64) (f64.abs (local.get 0)))
)

;; ── fadd's zero cases, read through 1/z so the sign is observable ──────────
;; ±0 + ∓0 = +0, in both orders.
(assert_return (invoke "add_sign" (f64.const 0) (f64.const -0)) (f64.const inf))
(assert_return (invoke "add_sign" (f64.const -0) (f64.const 0)) (f64.const inf))
;; ±0 + ±0 = that zero — so -0 + -0 keeps the sign, unlike the case above.
(assert_return (invoke "add_sign" (f64.const -0) (f64.const -0)) (f64.const -inf))
(assert_return (invoke "add_sign" (f64.const 0) (f64.const 0)) (f64.const inf))
;; ±q + ∓q = +0: exact cancellation is POSITIVE zero regardless of operand order.
(assert_return (invoke "add_sign" (f64.const 1.5) (f64.const -1.5)) (f64.const inf))
(assert_return (invoke "add_sign" (f64.const -1.5) (f64.const 1.5)) (f64.const inf))

;; z + ±0 returns z unchanged, including when z is negative.
(assert_return (invoke "add" (f64.const -3.25) (f64.const 0)) (f64.const -3.25))
(assert_return (invoke "add" (f64.const -3.25) (f64.const -0)) (f64.const -3.25))

;; ── fadd's infinity cases ─────────────────────────────────────────────────
(assert_return (invoke "add" (f64.const inf) (f64.const inf)) (f64.const inf))
(assert_return (invoke "add" (f64.const -inf) (f64.const -inf)) (f64.const -inf))
(assert_return (invoke "add" (f64.const inf) (f64.const 1)) (f64.const inf))
(assert_return (invoke "add" (f64.const -inf) (f64.const 1e308)) (f64.const -inf))
;; Opposite-sign infinities are the one NaN case in fadd.
(assert_return (invoke "add" (f64.const inf) (f64.const -inf)) (f64.const nan:canonical))
(assert_return (invoke "add" (f64.const -inf) (f64.const inf)) (f64.const nan:canonical))
;; The same shape in sub, mul and div.
(assert_return (invoke "sub" (f64.const inf) (f64.const inf)) (f64.const nan:canonical))
(assert_return (invoke "mul" (f64.const 0) (f64.const inf)) (f64.const nan:canonical))
(assert_return (invoke "mul" (f64.const -0) (f64.const inf)) (f64.const nan:canonical))
(assert_return (invoke "div" (f64.const 0) (f64.const 0)) (f64.const nan:canonical))
(assert_return (invoke "div" (f64.const inf) (f64.const inf)) (f64.const nan:canonical))

;; Division by zero is an INFINITY, not a trap — the sign is the xor of signs.
(assert_return (invoke "div" (f64.const 1) (f64.const 0)) (f64.const inf))
(assert_return (invoke "div" (f64.const 1) (f64.const -0)) (f64.const -inf))
(assert_return (invoke "div" (f64.const -1) (f64.const 0)) (f64.const -inf))
(assert_return (invoke "div" (f64.const -1) (f64.const -0)) (f64.const inf))

;; ── fmin / fmax: the zero rule is explicit because ±0 compare EQUAL ───────
;; fmin(±0, ∓0) = -0 in either order; fmax(±0, ∓0) = +0.
(assert_return (invoke "min_sign" (f64.const 0) (f64.const -0)) (f64.const -inf))
(assert_return (invoke "min_sign" (f64.const -0) (f64.const 0)) (f64.const -inf))
(assert_return (invoke "max_sign" (f64.const 0) (f64.const -0)) (f64.const inf))
(assert_return (invoke "max_sign" (f64.const -0) (f64.const 0)) (f64.const inf))

;; Infinity is handled before the ordinary comparison.
(assert_return (invoke "min" (f64.const -inf) (f64.const 0)) (f64.const -inf))
(assert_return (invoke "min" (f64.const inf) (f64.const 5)) (f64.const 5))
(assert_return (invoke "max" (f64.const inf) (f64.const 5)) (f64.const inf))
(assert_return (invoke "max" (f64.const -inf) (f64.const 5)) (f64.const 5))
(assert_return (invoke "min" (f64.const -inf) (f64.const inf)) (f64.const -inf))
(assert_return (invoke "max" (f64.const -inf) (f64.const inf)) (f64.const inf))

;; NaN wins over EVERYTHING in min/max — including infinity. This is where the
;; spec departs from the IEEE-754 minNum/maxNum recommendation, which would
;; return the non-NaN operand.
(assert_return (invoke "min" (f64.const nan) (f64.const 1)) (f64.const nan:canonical))
(assert_return (invoke "min" (f64.const 1) (f64.const nan)) (f64.const nan:canonical))
(assert_return (invoke "max" (f64.const nan) (f64.const 1)) (f64.const nan:canonical))
(assert_return (invoke "max" (f64.const 1) (f64.const nan)) (f64.const nan:canonical))
(assert_return (invoke "min" (f64.const nan) (f64.const -inf)) (f64.const nan:canonical))
(assert_return (invoke "max" (f64.const nan) (f64.const inf)) (f64.const nan:canonical))
(assert_return (invoke "min32" (f32.const nan) (f32.const 1)) (f32.const nan:canonical))
(assert_return (invoke "max32" (f32.const nan) (f32.const 1)) (f32.const nan:canonical))
(assert_return (invoke "min32" (f32.const -inf) (f32.const 0)) (f32.const -inf))

;; Ordinary comparisons still behave.
(assert_return (invoke "min" (f64.const -1) (f64.const 1)) (f64.const -1))
(assert_return (invoke "max" (f64.const -1) (f64.const 1)) (f64.const 1))

;; ── neg / abs / copysign are the three that do NOT follow the NaN rule ────
;; The spec exempts them: they act on the sign bit only, so a NaN keeps its
;; payload and the sign is fully determined rather than non-deterministic.
(assert_return (invoke "neg" (f64.const 0)) (f64.const -0))
(assert_return (invoke "neg" (f64.const -0)) (f64.const 0))
(assert_return (invoke "neg" (f64.const inf)) (f64.const -inf))
(assert_return (invoke "abs" (f64.const -0)) (f64.const 0))
(assert_return (invoke "abs" (f64.const -inf)) (f64.const inf))
(assert_return (invoke "copysign" (f64.const 3) (f64.const -1)) (f64.const -3))
(assert_return (invoke "copysign" (f64.const -3) (f64.const 1)) (f64.const 3))
(assert_return (invoke "copysign" (f64.const 3) (f64.const -0)) (f64.const -3))
(assert_return (invoke "copysign" (f64.const -0) (f64.const 1)) (f64.const 0))
(assert_return (invoke "copysign" (f64.const inf) (f64.const -1)) (f64.const -inf))
