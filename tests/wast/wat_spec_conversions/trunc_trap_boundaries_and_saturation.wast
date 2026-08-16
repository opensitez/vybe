;; vybe-test: wast/wat_spec_conversions/trunc_trap_boundaries_and_saturation
;; vybe-test-mode: run
;;
;; From the spec's `trunc` / `trunc_sat` / `convert` / `reinterpret` operators
;; (core/exec/numerics.rst, `aux-trunc` and `aux-sat`).
;;
;; `trunc` is partial: it is undefined for NaN, for either infinity, and for
;; any value whose truncation falls outside the target range. `trunc_sat` is
;; total and clamps instead. The interesting part is exactly WHERE the boundary
;; sits, because truncation happens BEFORE the range check:
;;
;;   trunc_u(-0.5) = 0        -- defined: truncates to 0, which is in range
;;   trunc_u(-1.0) = trap     -- truncates to -1, which is not
;;
;; so an implementation that range-checks the float before truncating rejects
;; -0.5 and is wrong. The signed boundary is the mirror image: the limit is
;; representable at the negative end (-2^31 exactly) but not at the positive
;; end (2^31 is one past the maximum), so the two ends are not symmetric.

(module
  (func (export "i32_s_f64") (param f64) (result i32) (i32.trunc_f64_s (local.get 0)))
  (func (export "i32_u_f64") (param f64) (result i32) (i32.trunc_f64_u (local.get 0)))
  (func (export "i32_s_f32") (param f32) (result i32) (i32.trunc_f32_s (local.get 0)))
  (func (export "i32_u_f32") (param f32) (result i32) (i32.trunc_f32_u (local.get 0)))
  (func (export "i64_s_f64") (param f64) (result i64) (i64.trunc_f64_s (local.get 0)))
  (func (export "i64_u_f64") (param f64) (result i64) (i64.trunc_f64_u (local.get 0)))

  (func (export "sat_i32_s") (param f64) (result i32) (i32.trunc_sat_f64_s (local.get 0)))
  (func (export "sat_i32_u") (param f64) (result i32) (i32.trunc_sat_f64_u (local.get 0)))
  (func (export "sat_i64_s") (param f64) (result i64) (i64.trunc_sat_f64_s (local.get 0)))
  (func (export "sat_i64_u") (param f64) (result i64) (i64.trunc_sat_f64_u (local.get 0)))

  (func (export "cvt_f64_i32_s") (param i32) (result f64) (f64.convert_i32_s (local.get 0)))
  (func (export "cvt_f64_i32_u") (param i32) (result f64) (f64.convert_i32_u (local.get 0)))
  (func (export "cvt_f64_i64_u") (param i64) (result f64) (f64.convert_i64_u (local.get 0)))
  (func (export "cvt_f32_i32_s") (param i32) (result f32) (f32.convert_i32_s (local.get 0)))

  (func (export "demote") (param f64) (result f32) (f32.demote_f64 (local.get 0)))
  (func (export "promote") (param f32) (result f64) (f64.promote_f32 (local.get 0)))
  (func (export "bits_f32") (param f32) (result i32) (i32.reinterpret_f32 (local.get 0)))
  (func (export "bits_f64") (param f64) (result i64) (i64.reinterpret_f64 (local.get 0)))
  (func (export "from_bits_f32") (param i32) (result f32) (f32.reinterpret_i32 (local.get 0)))
  (func (export "from_bits_f64") (param i64) (result f64) (f64.reinterpret_i64 (local.get 0)))
)

;; ── Truncation happens BEFORE the range check ─────────────────────────────
;; Everything in (-1, 0] truncates to 0 and is therefore in range for the
;; UNSIGNED conversion, negative sign notwithstanding.
(assert_return (invoke "i32_u_f64" (f64.const -0.5)) (i32.const 0))
(assert_return (invoke "i32_u_f64" (f64.const -0.9999999999)) (i32.const 0))
(assert_return (invoke "i32_u_f64" (f64.const -0.0)) (i32.const 0))
;; -1.0 truncates to -1, which is not.
(assert_trap (invoke "i32_u_f64" (f64.const -1.0)) "integer overflow")
(assert_trap (invoke "i32_u_f64" (f64.const -1.5)) "integer overflow")

;; The signed boundary is asymmetric: -2^31 is representable, +2^31 is not.
(assert_return (invoke "i32_s_f64" (f64.const -2147483648.0)) (i32.const -2147483648))
(assert_return (invoke "i32_s_f64" (f64.const 2147483647.0)) (i32.const 2147483647))
(assert_trap (invoke "i32_s_f64" (f64.const 2147483648.0)) "integer overflow")
(assert_trap (invoke "i32_s_f64" (f64.const -2147483649.0)) "integer overflow")
;; ...but a fractional part inside the boundary truncates toward zero and is fine.
(assert_return (invoke "i32_s_f64" (f64.const -2147483648.9)) (i32.const -2147483648))
(assert_return (invoke "i32_s_f64" (f64.const 2147483647.9)) (i32.const 2147483647))

;; Unsigned upper boundary.
(assert_return (invoke "i32_u_f64" (f64.const 4294967295.0)) (i32.const -1))
(assert_return (invoke "i32_u_f64" (f64.const 4294967295.9)) (i32.const -1))
(assert_trap (invoke "i32_u_f64" (f64.const 4294967296.0)) "integer overflow")

;; NaN and both infinities are undefined for every `trunc`.
(assert_trap (invoke "i32_s_f64" (f64.const nan)) "invalid conversion to integer")
(assert_trap (invoke "i32_u_f64" (f64.const nan)) "invalid conversion to integer")
(assert_trap (invoke "i32_s_f64" (f64.const inf)) "integer overflow")
(assert_trap (invoke "i32_s_f64" (f64.const -inf)) "integer overflow")
(assert_trap (invoke "i64_s_f64" (f64.const nan)) "invalid conversion to integer")
(assert_trap (invoke "i64_u_f64" (f64.const -1.0)) "integer overflow")
(assert_return (invoke "i64_u_f64" (f64.const -0.5)) (i64.const 0))

;; f32 sources: 2^31 is exactly representable in f32 and is out of range.
(assert_return (invoke "i32_s_f32" (f32.const -2147483648.0)) (i32.const -2147483648))
(assert_trap (invoke "i32_s_f32" (f32.const 2147483648.0)) "integer overflow")
(assert_return (invoke "i32_u_f32" (f32.const -0.5)) (i32.const 0))
(assert_trap (invoke "i32_u_f32" (f32.const -1.0)) "integer overflow")

;; Truncation is toward zero on both sides, never floor.
(assert_return (invoke "i32_s_f64" (f64.const 1.9)) (i32.const 1))
(assert_return (invoke "i32_s_f64" (f64.const -1.9)) (i32.const -1))
(assert_return (invoke "i32_s_f64" (f64.const -0.9)) (i32.const 0))

;; ── trunc_sat is total: the same inputs clamp instead of trapping ─────────
(assert_return (invoke "sat_i32_s" (f64.const nan)) (i32.const 0))
(assert_return (invoke "sat_i32_s" (f64.const -nan)) (i32.const 0))
(assert_return (invoke "sat_i32_s" (f64.const inf)) (i32.const 2147483647))
(assert_return (invoke "sat_i32_s" (f64.const -inf)) (i32.const -2147483648))
(assert_return (invoke "sat_i32_s" (f64.const 2147483648.0)) (i32.const 2147483647))
(assert_return (invoke "sat_i32_s" (f64.const -2147483649.0)) (i32.const -2147483648))
(assert_return (invoke "sat_i32_u" (f64.const nan)) (i32.const 0))
(assert_return (invoke "sat_i32_u" (f64.const -1.0)) (i32.const 0))
(assert_return (invoke "sat_i32_u" (f64.const -inf)) (i32.const 0))
(assert_return (invoke "sat_i32_u" (f64.const inf)) (i32.const -1))
(assert_return (invoke "sat_i32_u" (f64.const 4294967296.0)) (i32.const -1))
(assert_return (invoke "sat_i64_s" (f64.const inf)) (i64.const 9223372036854775807))
(assert_return (invoke "sat_i64_s" (f64.const -inf)) (i64.const -9223372036854775808))
(assert_return (invoke "sat_i64_u" (f64.const -1.0)) (i64.const 0))
;; In range, sat and non-sat agree exactly.
(assert_return (invoke "sat_i32_s" (f64.const -1.9)) (i32.const -1))
(assert_return (invoke "sat_i32_u" (f64.const -0.5)) (i32.const 0))

;; ── convert: the sign of the source decides, and f32 rounds ──────────────
(assert_return (invoke "cvt_f64_i32_s" (i32.const -1)) (f64.const -1))
(assert_return (invoke "cvt_f64_i32_u" (i32.const -1)) (f64.const 4294967295))
(assert_return (invoke "cvt_f64_i32_u" (i32.const -2147483648)) (f64.const 2147483648))
(assert_return (invoke "cvt_f64_i64_u" (i64.const -1)) (f64.const 18446744073709551616))
;; 16777217 is the first integer f32 cannot represent; it rounds to even.
(assert_return (invoke "cvt_f32_i32_s" (i32.const 16777217)) (f32.const 16777216))
(assert_return (invoke "cvt_f32_i32_s" (i32.const 16777216)) (f32.const 16777216))

;; ── demote / promote ─────────────────────────────────────────────────────
;; Overflow on demote is an infinity, not a trap.
(assert_return (invoke "demote" (f64.const 1e300)) (f32.const inf))
(assert_return (invoke "demote" (f64.const -1e300)) (f32.const -inf))
;; Underflow reaches zero and keeps its sign.
(assert_return (invoke "demote" (f64.const -1e-300)) (f32.const -0))
(assert_return (invoke "demote" (f64.const 1.5)) (f32.const 1.5))
;; promote is exact — every f32 is an f64.
(assert_return (invoke "promote" (f32.const 1.5)) (f64.const 1.5))
(assert_return (invoke "promote" (f32.const -0)) (f64.const -0))
(assert_return (invoke "promote" (f32.const inf)) (f64.const inf))

;; ── reinterpret moves BITS, not values ───────────────────────────────────
(assert_return (invoke "bits_f32" (f32.const 1)) (i32.const 1065353216))
(assert_return (invoke "bits_f32" (f32.const -0)) (i32.const -2147483648))
(assert_return (invoke "bits_f32" (f32.const 0)) (i32.const 0))
(assert_return (invoke "from_bits_f32" (i32.const 1065353216)) (f32.const 1))
(assert_return (invoke "from_bits_f32" (i32.const -2147483648)) (f32.const -0))
(assert_return (invoke "bits_f64" (f64.const 1)) (i64.const 4607182418800017408))
(assert_return (invoke "bits_f64" (f64.const -0)) (i64.const -9223372036854775808))
(assert_return (invoke "from_bits_f64" (i64.const 4607182418800017408)) (f64.const 1))
;; A NaN bit pattern survives the round trip as a NaN.
(assert_return (invoke "from_bits_f64" (i64.const 9218868437227405312)) (f64.const inf))
