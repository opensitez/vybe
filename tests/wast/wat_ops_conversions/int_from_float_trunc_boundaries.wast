;; vybe-test: wast/wat_ops_conversions/int_from_float_trunc_boundaries
;; origin: coverage gap vs crates/vybe_runtime/tests/numeric_conversions_test.rs
;; vybe-test-mode: run
;;
;; The trapping float→int truncations, at their RANGE BOUNDARIES.
;;
;; `trunc` rounds toward zero, so the valid domain of the unsigned forms is
;; the OPEN interval (-1, 2^N) — every value strictly greater than -1
;; truncates to 0 and is legal, including tiny negatives and -0.0. Only -1.0
;; and below trap. That distinction is the whole content of these
;; instructions and it is invisible to any test that checks a single
;; negative: `-1.0` traps, `-0x1p-149` must return 0, and a test picking only
;; the former concludes "negatives trap" and passes an implementation that
;; traps on both.
;;
;; (`crates/vybe_runtime/tests/numeric_conversions_test.rs:112` is exactly
;; that test — `i32_trunc_f32_u_neg_traps` checks -1.0 alone.)
;;
;; Spec-format so `wasmtime wast` arbitrates the boundary rather than us.

(module
  (func (export "i32.trunc_f32_u") (param f32) (result i32) (i32.trunc_f32_u (local.get 0)))
  (func (export "i32.trunc_f64_u") (param f64) (result i32) (i32.trunc_f64_u (local.get 0)))
  (func (export "i64.trunc_f32_s") (param f32) (result i64) (i64.trunc_f32_s (local.get 0)))
  (func (export "i64.trunc_f32_u") (param f32) (result i64) (i64.trunc_f32_u (local.get 0)))
  (func (export "i64.trunc_f64_u") (param f64) (result i64) (i64.trunc_f64_u (local.get 0)))
)

;; ── inside (-1, 0]: legal, truncates to zero ────────────────────────────
(assert_return (invoke "i32.trunc_f32_u" (f32.const -0x1p-149)) (i32.const 0))
(assert_return (invoke "i32.trunc_f32_u" (f32.const -0.0)) (i32.const 0))
(assert_return (invoke "i32.trunc_f32_u" (f32.const -0.9)) (i32.const 0))
(assert_return (invoke "i32.trunc_f64_u" (f64.const -0x1p-1074)) (i32.const 0))
(assert_return (invoke "i32.trunc_f64_u" (f64.const -0.9)) (i32.const 0))
(assert_return (invoke "i64.trunc_f32_u" (f32.const -0x1p-149)) (i64.const 0))
(assert_return (invoke "i64.trunc_f64_u" (f64.const -0.9)) (i64.const 0))

;; ── at and below -1.0: traps ────────────────────────────────────────────
;; "integer overflow", not "invalid conversion to integer" — the value is
;; OUT OF RANGE, which is a different trap from the un-convertible NaN case
;; below (`conversions.wast:101` vs `:104`).
(assert_trap (invoke "i32.trunc_f32_u" (f32.const -1.0)) "integer overflow")
(assert_trap (invoke "i32.trunc_f64_u" (f64.const -1.0)) "integer overflow")
(assert_trap (invoke "i64.trunc_f32_u" (f32.const -1.0)) "integer overflow")
(assert_trap (invoke "i64.trunc_f64_u" (f64.const -1.0)) "integer overflow")

;; ── ordinary values, and truncation toward zero (never rounding) ────────
(assert_return (invoke "i32.trunc_f32_u" (f32.const 1.9)) (i32.const 1))
(assert_return (invoke "i32.trunc_f64_u" (f64.const 1.9)) (i32.const 1))
(assert_return (invoke "i64.trunc_f32_s" (f32.const -1.9)) (i64.const -1))
(assert_return (invoke "i64.trunc_f32_s" (f32.const 1.9)) (i64.const 1))

;; ── the top of the unsigned range ───────────────────────────────────────
;; 0xffffff00 is the largest f32-representable value below 2^32.
(assert_return (invoke "i32.trunc_f32_u" (f32.const 0xffffff00)) (i32.const -256))
(assert_trap  (invoke "i32.trunc_f32_u" (f32.const 4294967296.0)) "integer overflow")
(assert_return (invoke "i32.trunc_f64_u" (f64.const 4294967295.0)) (i32.const -1))
(assert_trap  (invoke "i32.trunc_f64_u" (f64.const 4294967296.0)) "integer overflow")

;; ── signed i64 boundaries ───────────────────────────────────────────────
(assert_return (invoke "i64.trunc_f32_s" (f32.const -9223372036854775808.0)) (i64.const -9223372036854775808))
(assert_trap  (invoke "i64.trunc_f32_s" (f32.const 9223372036854775808.0)) "integer overflow")

;; ── NaN and infinity are never convertible ──────────────────────────────
;; NaN is UNCONVERTIBLE (a different trap); infinities are merely OUT OF RANGE.
(assert_trap (invoke "i32.trunc_f32_u" (f32.const nan)) "invalid conversion to integer")
(assert_trap (invoke "i32.trunc_f32_u" (f32.const nan:0x200000)) "invalid conversion to integer")
(assert_trap (invoke "i32.trunc_f32_u" (f32.const inf)) "integer overflow")
(assert_trap (invoke "i64.trunc_f32_s" (f32.const -inf)) "integer overflow")
