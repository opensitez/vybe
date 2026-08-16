;; vybe-test: wast/wat_ops_conversions/trunc_sat_never_traps
;; origin: coverage gap — four of the eight trunc_sat ops occurred ONCE in the run corpus
;; vybe-test-mode: run
;;
;; `trunc_sat` is `trunc` with the trap replaced by SATURATION. Same rounding,
;; same domain — different behaviour outside it, and that difference is the
;; entire reason the instruction exists.
;;
;; So the only assertions that distinguish the two families are the ones OUT of
;; range, and there are exactly three kinds:
;;
;;   * above the maximum  → the maximum (not wraparound, not the trap)
;;   * below the minimum  → the minimum
;;   * NaN                → ZERO, for both signednesses — not the minimum, which
;;                          is what a naive clamp of an unordered value gives
;;
;; A test that only checks in-range values is checking `trunc`, and passes
;; against an implementation where `trunc_sat` still traps.
;;
;; The unsigned forms saturate to 0 at the bottom, and their valid domain is
;; the OPEN interval (-1, 2^N) — so -0.9 is 0 by TRUNCATION and -1.0 is 0 by
;; SATURATION, two different mechanisms reaching the same answer.

(module
  (func (export "i32_sat_f32_s") (param f32) (result i32) (i32.trunc_sat_f32_s (local.get 0)))
  (func (export "i32_sat_f32_u") (param f32) (result i32) (i32.trunc_sat_f32_u (local.get 0)))
  (func (export "i32_sat_f64_s") (param f64) (result i32) (i32.trunc_sat_f64_s (local.get 0)))
  (func (export "i32_sat_f64_u") (param f64) (result i32) (i32.trunc_sat_f64_u (local.get 0)))
  (func (export "i64_sat_f32_s") (param f32) (result i64) (i64.trunc_sat_f32_s (local.get 0)))
  (func (export "i64_sat_f32_u") (param f32) (result i64) (i64.trunc_sat_f32_u (local.get 0)))
  (func (export "i64_sat_f64_s") (param f64) (result i64) (i64.trunc_sat_f64_s (local.get 0)))
  (func (export "i64_sat_f64_u") (param f64) (result i64) (i64.trunc_sat_f64_u (local.get 0)))
)

;; ── in range: identical to the trapping forms ───────────────────────────
(assert_return (invoke "i32_sat_f32_s" (f32.const 1.9)) (i32.const 1))
(assert_return (invoke "i32_sat_f32_s" (f32.const -1.9)) (i32.const -1))
(assert_return (invoke "i32_sat_f64_s" (f64.const -1.9)) (i32.const -1))
(assert_return (invoke "i64_sat_f64_s" (f64.const 1.9)) (i64.const 1))
;; Truncation toward zero, NOT saturation, is what makes these 0.
(assert_return (invoke "i32_sat_f32_u" (f32.const -0.9)) (i32.const 0))
(assert_return (invoke "i32_sat_f64_u" (f64.const -0.9)) (i32.const 0))

;; ── NaN is ZERO, for every form ─────────────────────────────────────────
;; A clamp written as `min(max(v, LO), HI)` on an unordered value returns LO,
;; so the signed forms would give i32.min here rather than 0.
(assert_return (invoke "i32_sat_f32_s" (f32.const nan)) (i32.const 0))
(assert_return (invoke "i32_sat_f32_s" (f32.const -nan)) (i32.const 0))
(assert_return (invoke "i32_sat_f32_s" (f32.const nan:0x200000)) (i32.const 0))
(assert_return (invoke "i32_sat_f32_u" (f32.const nan)) (i32.const 0))
(assert_return (invoke "i32_sat_f64_s" (f64.const nan)) (i32.const 0))
(assert_return (invoke "i32_sat_f64_u" (f64.const nan)) (i32.const 0))
(assert_return (invoke "i64_sat_f32_s" (f32.const nan)) (i64.const 0))
(assert_return (invoke "i64_sat_f32_u" (f32.const nan)) (i64.const 0))
(assert_return (invoke "i64_sat_f64_s" (f64.const nan)) (i64.const 0))
(assert_return (invoke "i64_sat_f64_u" (f64.const nan)) (i64.const 0))

;; ── out of range saturates instead of trapping ──────────────────────────
(assert_return (invoke "i32_sat_f32_s" (f32.const inf)) (i32.const 0x7fffffff))
(assert_return (invoke "i32_sat_f32_s" (f32.const -inf)) (i32.const 0x80000000))
(assert_return (invoke "i32_sat_f32_s" (f32.const 1e30)) (i32.const 0x7fffffff))
(assert_return (invoke "i32_sat_f32_s" (f32.const -1e30)) (i32.const 0x80000000))
(assert_return (invoke "i32_sat_f64_s" (f64.const 1e30)) (i32.const 0x7fffffff))
(assert_return (invoke "i32_sat_f64_s" (f64.const -1e30)) (i32.const 0x80000000))
;; Unsigned saturates to 0 at the bottom, not to the signed minimum.
(assert_return (invoke "i32_sat_f32_u" (f32.const -inf)) (i32.const 0))
(assert_return (invoke "i32_sat_f32_u" (f32.const -1.0)) (i32.const 0))
(assert_return (invoke "i32_sat_f32_u" (f32.const -1e30)) (i32.const 0))
(assert_return (invoke "i32_sat_f32_u" (f32.const inf)) (i32.const 0xffffffff))
(assert_return (invoke "i32_sat_f64_u" (f64.const -1.0)) (i32.const 0))
(assert_return (invoke "i32_sat_f64_u" (f64.const 1e30)) (i32.const 0xffffffff))
(assert_return (invoke "i64_sat_f32_s" (f32.const inf)) (i64.const 0x7fffffffffffffff))
(assert_return (invoke "i64_sat_f32_s" (f32.const -inf)) (i64.const 0x8000000000000000))
(assert_return (invoke "i64_sat_f32_u" (f32.const -1.0)) (i64.const 0))
(assert_return (invoke "i64_sat_f32_u" (f32.const inf)) (i64.const 0xffffffffffffffff))
(assert_return (invoke "i64_sat_f64_s" (f64.const 1e300)) (i64.const 0x7fffffffffffffff))
(assert_return (invoke "i64_sat_f64_u" (f64.const 1e300)) (i64.const 0xffffffffffffffff))

;; ── the exact boundary: last representable in, first out ────────────────
;; 2^31 is out of range for signed i32; the f32 below it is the largest in.
(assert_return (invoke "i32_sat_f32_s" (f32.const 2147483648.0)) (i32.const 0x7fffffff))
(assert_return (invoke "i32_sat_f32_s" (f32.const 0x1.fffffep+30)) (i32.const 2147483520))
(assert_return (invoke "i32_sat_f32_s" (f32.const -2147483648.0)) (i32.const 0x80000000))
;; 2^32 is out for unsigned; 0xffffff00 is the largest f32 below it.
(assert_return (invoke "i32_sat_f32_u" (f32.const 4294967296.0)) (i32.const 0xffffffff))
(assert_return (invoke "i32_sat_f32_u" (f32.const 0xffffff00)) (i32.const 0xffffff00))
(assert_return (invoke "i64_sat_f64_s" (f64.const 9223372036854775808.0)) (i64.const 0x7fffffffffffffff))
(assert_return (invoke "i64_sat_f64_s" (f64.const -9223372036854775808.0)) (i64.const 0x8000000000000000))
