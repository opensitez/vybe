;; vybe-test: wast/wat_ops_f32/f32_arithmetic_zeros_and_nan
;; origin: coverage gap vs tests/wast/wat_f32_arithmetic (22 files, ONE assertion each)
;; vybe-test-mode: run
;;
;; f32 add/sub/mul/div/neg/sqrt/min/max at the points where they are not
;; ordinary arithmetic: signed zeros, infinities, NaN propagation, and f32
;; ROUNDING.
;;
;; `tests/wast/wat_f32_arithmetic` spends 22 files on these ops and asserts one
;; ordinary value each (`-3.14 abs` → `3.14`). Every one of them passes on an
;; implementation that computes in DOUBLE precision and never propagates a NaN
;; — which is the actual risk here, since this VM carries f64 alongside f32.
;; The three properties below are what separate a real f32 from a narrowed f64:
;;
;;   * ROUNDING happens at 24 bits, not 53. `1 + 2^-24` is exactly 1.0 in f32
;;     and is NOT 1.0 in f64, so a double-precision path fails only here.
;;   * SIGNED ZERO survives: `-0.0` and `0.0` compare equal, so `f32.eq` cannot
;;     see the difference — only `copysign`/`div` expose it.
;;   * min/max are NOT `a < b ? a : b`. They return NaN if either operand is
;;     NaN, and they order `-0.0` below `+0.0` where `<` calls them equal.
;;
;; Spec-format so `wasmtime wast` arbitrates every expectation.

(module
  (func (export "add") (param f32 f32) (result f32) (f32.add (local.get 0) (local.get 1)))
  (func (export "sub") (param f32 f32) (result f32) (f32.sub (local.get 0) (local.get 1)))
  (func (export "mul") (param f32 f32) (result f32) (f32.mul (local.get 0) (local.get 1)))
  (func (export "div") (param f32 f32) (result f32) (f32.div (local.get 0) (local.get 1)))
  (func (export "neg") (param f32) (result f32) (f32.neg (local.get 0)))
  (func (export "sqrt") (param f32) (result f32) (f32.sqrt (local.get 0)))
  (func (export "min") (param f32 f32) (result f32) (f32.min (local.get 0) (local.get 1)))
  (func (export "max") (param f32 f32) (result f32) (f32.max (local.get 0) (local.get 1)))
  ;; Bit-returning forms. A signed zero and a NaN sign are invisible to
  ;; `f32.eq`, and an assertion cannot nest one invoke inside another, so each
  ;; case that needs the raw pattern gets its own export.
  (func (export "add_bits") (param f32 f32) (result i32) (i32.reinterpret_f32 (f32.add (local.get 0) (local.get 1))))
  (func (export "mul_bits") (param f32 f32) (result i32) (i32.reinterpret_f32 (f32.mul (local.get 0) (local.get 1))))
  (func (export "neg_bits") (param f32) (result i32) (i32.reinterpret_f32 (f32.neg (local.get 0))))
  (func (export "sqrt_bits") (param f32) (result i32) (i32.reinterpret_f32 (f32.sqrt (local.get 0))))
  (func (export "min_bits") (param f32 f32) (result i32) (i32.reinterpret_f32 (f32.min (local.get 0) (local.get 1))))
  (func (export "max_bits") (param f32 f32) (result i32) (i32.reinterpret_f32 (f32.max (local.get 0) (local.get 1))))
)

;; ── rounding is at 24 bits, not 53 ──────────────────────────────────────
;; 2^-24 is below the f32 gap at 1.0, so it rounds away entirely. In f64 the
;; same sum is representable — this is the assertion a double path fails.
(assert_return (invoke "add" (f32.const 1.0) (f32.const 0x1p-24)) (f32.const 1.0))
;; 2^-23 IS the gap at 1.0, so this one survives.
(assert_return (invoke "add" (f32.const 1.0) (f32.const 0x1p-23)) (f32.const 0x1.000002p+0))
;; Ties round to even: 1 + 1.5·2^-24 lands on the next representable value.
(assert_return (invoke "add" (f32.const 1.0) (f32.const 0x1.8p-24)) (f32.const 0x1.000002p+0))

;; ── signed zero ─────────────────────────────────────────────────────────
;; x + -x is +0.0, not -0.0 (round-to-nearest). Only the bit pattern shows it.
(assert_return (invoke "add_bits" (f32.const 1.0) (f32.const -1.0)) (i32.const 0))
(assert_return (invoke "neg_bits" (f32.const 0.0)) (i32.const 0x80000000))
(assert_return (invoke "mul_bits" (f32.const -1.0) (f32.const 0.0)) (i32.const 0x80000000))
;; -0.0 == +0.0 is TRUE, which is why the comparisons above use bits.
(assert_return (invoke "sub" (f32.const 0.0) (f32.const 0.0)) (f32.const 0.0))
;; sqrt(-0.0) is -0.0 — the one negative input sqrt does not turn into NaN.
(assert_return (invoke "sqrt_bits" (f32.const -0.0)) (i32.const 0x80000000))

;; ── division: zeros and infinities ──────────────────────────────────────
(assert_return (invoke "div" (f32.const 1.0) (f32.const 0.0)) (f32.const inf))
(assert_return (invoke "div" (f32.const 1.0) (f32.const -0.0)) (f32.const -inf))
(assert_return (invoke "div" (f32.const -1.0) (f32.const 0.0)) (f32.const -inf))
(assert_return (invoke "div" (f32.const 0.0) (f32.const 0.0)) (f32.const nan:canonical))
(assert_return (invoke "div" (f32.const inf) (f32.const inf)) (f32.const nan:canonical))
(assert_return (invoke "div" (f32.const 1.0) (f32.const inf)) (f32.const 0.0))
;; inf - inf and 0 * inf are the two other NaN-producing arithmetic cases.
(assert_return (invoke "sub" (f32.const inf) (f32.const inf)) (f32.const nan:canonical))
(assert_return (invoke "mul" (f32.const 0.0) (f32.const inf)) (f32.const nan:canonical))
(assert_return (invoke "add" (f32.const inf) (f32.const 1.0)) (f32.const inf))

;; ── overflow to infinity, underflow to zero ─────────────────────────────
;; The largest finite f32 doubled is not representable, so it becomes inf
;; rather than the f64 value that would fit.
(assert_return (invoke "mul" (f32.const 0x1.fffffep+127) (f32.const 2.0)) (f32.const inf))
(assert_return (invoke "mul" (f32.const 0x1p-149) (f32.const 0.5)) (f32.const 0.0))
;; Subnormals are still arithmetic, not flushed to zero.
(assert_return (invoke "add" (f32.const 0x1p-149) (f32.const 0x1p-149)) (f32.const 0x1p-148))

;; ── sqrt ────────────────────────────────────────────────────────────────
(assert_return (invoke "sqrt" (f32.const 4.0)) (f32.const 2.0))
(assert_return (invoke "sqrt" (f32.const -1.0)) (f32.const nan:canonical))
(assert_return (invoke "sqrt" (f32.const inf)) (f32.const inf))

;; ── min / max are not `<` ───────────────────────────────────────────────
;; NaN wins over any operand, in EITHER position.
(assert_return (invoke "min" (f32.const nan) (f32.const 1.0)) (f32.const nan:canonical))
(assert_return (invoke "min" (f32.const 1.0) (f32.const nan)) (f32.const nan:canonical))
(assert_return (invoke "max" (f32.const nan) (f32.const 1.0)) (f32.const nan:canonical))
;; -0.0 and +0.0 compare EQUAL, so a `<`-based implementation returns whichever
;; operand came first. min must pick -0.0 and max +0.0 regardless of order.
(assert_return (invoke "min_bits" (f32.const -0.0) (f32.const 0.0)) (i32.const 0x80000000))
(assert_return (invoke "min_bits" (f32.const 0.0) (f32.const -0.0)) (i32.const 0x80000000))
(assert_return (invoke "max_bits" (f32.const -0.0) (f32.const 0.0)) (i32.const 0))
(assert_return (invoke "max_bits" (f32.const 0.0) (f32.const -0.0)) (i32.const 0))
(assert_return (invoke "min" (f32.const inf) (f32.const 1.0)) (f32.const 1.0))
(assert_return (invoke "max" (f32.const -inf) (f32.const -1.0)) (f32.const -1.0))

;; ── NaN propagates through arithmetic ───────────────────────────────────
(assert_return (invoke "add" (f32.const nan) (f32.const 1.0)) (f32.const nan:canonical))
(assert_return (invoke "mul" (f32.const nan) (f32.const 0.0)) (f32.const nan:canonical))
;; neg flips the SIGN of a NaN rather than quieting or canonicalising it.
(assert_return (invoke "neg_bits" (f32.const nan)) (i32.const 0xffc00000))
(assert_return (invoke "neg_bits" (f32.const -nan)) (i32.const 0x7fc00000))
