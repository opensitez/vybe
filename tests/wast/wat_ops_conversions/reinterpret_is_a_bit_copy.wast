;; vybe-test: wast/wat_ops_conversions/reinterpret_is_a_bit_copy
;; origin: coverage gap — all four reinterpret ops occurred ONCE in the run corpus
;; vybe-test-mode: run
;;
;; The four `reinterpret` ops are a PURE BIT COPY. They are the only way to
;; observe a float's exact encoding, and the only float ops with no numeric
;; behaviour at all — no rounding, no quieting, no canonicalisation.
;;
;; That makes them the ops most easily broken by an implementation that routes
;; values through a wider type. Widening an f32 to f64 and narrowing back is
;; value-preserving for every FINITE float, so such a bug is invisible
;; everywhere except here:
;;
;;   * a SIGNALLING NaN gets quieted by the hardware widening (x86 `cvtss2sd`
;;     sets the quiet bit), so `0x7fa00000` returns as `0x7fc00000`;
;;   * a NaN PAYLOAD wider than 22 bits cannot survive the narrowing back.
;;
;; Both are exact-bit facts, so every assertion here is written as the integer
;; pattern rather than as a float.

(module
  (func (export "i32_of_f32") (param f32) (result i32) (i32.reinterpret_f32 (local.get 0)))
  (func (export "f32_of_i32") (param i32) (result f32) (f32.reinterpret_i32 (local.get 0)))
  (func (export "i64_of_f64") (param f64) (result i64) (i64.reinterpret_f64 (local.get 0)))
  (func (export "f64_of_i64") (param i64) (result f64) (f64.reinterpret_i64 (local.get 0)))
  ;; Round-trip in one call: bits → float → bits must be the identity for EVERY
  ;; pattern, including the ones that are not numbers.
  (func (export "rt32") (param i32) (result i32)
    (i32.reinterpret_f32 (f32.reinterpret_i32 (local.get 0))))
  (func (export "rt64") (param i64) (result i64)
    (i64.reinterpret_f64 (f64.reinterpret_i64 (local.get 0))))
)

;; ── ordinary values, both directions ────────────────────────────────────
(assert_return (invoke "i32_of_f32" (f32.const 1.0)) (i32.const 0x3f800000))
(assert_return (invoke "i32_of_f32" (f32.const -1.0)) (i32.const 0xbf800000))
(assert_return (invoke "f32_of_i32" (i32.const 0x3f800000)) (f32.const 1.0))
(assert_return (invoke "i64_of_f64" (f64.const 1.0)) (i64.const 0x3ff0000000000000))
(assert_return (invoke "f64_of_i64" (i64.const 0x3ff0000000000000)) (f64.const 1.0))

;; ── signed zero: the sign bit is data, not a value ──────────────────────
(assert_return (invoke "i32_of_f32" (f32.const 0.0)) (i32.const 0))
(assert_return (invoke "i32_of_f32" (f32.const -0.0)) (i32.const 0x80000000))
(assert_return (invoke "i64_of_f64" (f64.const -0.0)) (i64.const 0x8000000000000000))

;; ── infinities ──────────────────────────────────────────────────────────
(assert_return (invoke "i32_of_f32" (f32.const inf)) (i32.const 0x7f800000))
(assert_return (invoke "i32_of_f32" (f32.const -inf)) (i32.const 0xff800000))
(assert_return (invoke "i64_of_f64" (f64.const inf)) (i64.const 0x7ff0000000000000))

;; ── subnormals and the extremes of the finite range ─────────────────────
(assert_return (invoke "i32_of_f32" (f32.const 0x1p-149)) (i32.const 1))
(assert_return (invoke "i32_of_f32" (f32.const 0x1.fffffep+127)) (i32.const 0x7f7fffff))
(assert_return (invoke "f32_of_i32" (i32.const 1)) (f32.const 0x1p-149))
(assert_return (invoke "i64_of_f64" (f64.const 0x1p-1074)) (i64.const 1))

;; ── NaN sign and payload survive exactly ────────────────────────────────
(assert_return (invoke "i32_of_f32" (f32.const nan)) (i32.const 0x7fc00000))
(assert_return (invoke "i32_of_f32" (f32.const -nan)) (i32.const 0xffc00000))
(assert_return (invoke "i32_of_f32" (f32.const nan:0x7fffff)) (i32.const 0x7fffffff))
(assert_return (invoke "i32_of_f32" (f32.const -nan:0x7fffff)) (i32.const 0xffffffff))
;; SIGNALLING NaN — quiet bit CLEAR. This is the pattern a widen/narrow
;; round-trip destroys.
(assert_return (invoke "i32_of_f32" (f32.const nan:0x200000)) (i32.const 0x7fa00000))
(assert_return (invoke "i64_of_f64" (f64.const nan)) (i64.const 0x7ff8000000000000))
(assert_return (invoke "i64_of_f64" (f64.const -nan)) (i64.const 0xfff8000000000000))
(assert_return (invoke "i64_of_f64" (f64.const nan:0x4000000000000)) (i64.const 0x7ff4000000000000))

;; ── round-trip is the identity for every pattern ────────────────────────
(assert_return (invoke "rt32" (i32.const 0x7fa00000)) (i32.const 0x7fa00000))
(assert_return (invoke "rt32" (i32.const 0xffc00000)) (i32.const 0xffc00000))
(assert_return (invoke "rt32" (i32.const 0x00000001)) (i32.const 0x00000001))
(assert_return (invoke "rt32" (i32.const 0xffffffff)) (i32.const 0xffffffff))
(assert_return (invoke "rt64" (i64.const 0x7ff4000000000000)) (i64.const 0x7ff4000000000000))
(assert_return (invoke "rt64" (i64.const 0xffffffffffffffff)) (i64.const 0xffffffffffffffff))
