;; vybe-test: wast/wat_ops_conversions/int_to_float_convert_rounding
;; origin: coverage gap — six of the eight convert_* ops occurred ONCE in the run corpus
;; vybe-test-mode: run
;;
;; `convert` is int → float, and unlike `trunc` it never traps: every integer
;; has a float image. What it does instead is ROUND, and that is the whole
;; content of the instruction.
;;
;; Two facts no "does it convert 3 to 3.0" test reaches:
;;
;;   * SIGNEDNESS is in the opcode, not the operand. `-1` as i32 is
;;     4294967295 unsigned, so `f32.convert_i32_s(-1)` is -1.0 while
;;     `f32.convert_i32_u(-1)` is 4294967295.0 — and since f32 has only 24
;;     mantissa bits, that value is not representable and rounds UP to
;;     4294967296.0, a number larger than any u32.
;;   * ROUNDING is to nearest, ties to even — not truncation. An i64 above
;;     2^53 does not fit an f64 mantissa, so `f64.convert_i64_s` is lossy too;
;;     only `f64.convert_i32_*` is exact for its whole domain.

(module
  (func (export "f32.convert_i32_s") (param i32) (result f32) (f32.convert_i32_s (local.get 0)))
  (func (export "f32.convert_i32_u") (param i32) (result f32) (f32.convert_i32_u (local.get 0)))
  (func (export "f32.convert_i64_s") (param i64) (result f32) (f32.convert_i64_s (local.get 0)))
  (func (export "f32.convert_i64_u") (param i64) (result f32) (f32.convert_i64_u (local.get 0)))
  (func (export "f64.convert_i32_s") (param i32) (result f64) (f64.convert_i32_s (local.get 0)))
  (func (export "f64.convert_i32_u") (param i32) (result f64) (f64.convert_i32_u (local.get 0)))
  (func (export "f64.convert_i64_s") (param i64) (result f64) (f64.convert_i64_s (local.get 0)))
  (func (export "f64.convert_i64_u") (param i64) (result f64) (f64.convert_i64_u (local.get 0)))
)

;; ── signedness lives in the OPCODE ──────────────────────────────────────
(assert_return (invoke "f32.convert_i32_s" (i32.const -1)) (f32.const -1.0))
(assert_return (invoke "f32.convert_i32_u" (i32.const -1)) (f32.const 4294967296.0))
(assert_return (invoke "f64.convert_i32_s" (i32.const -1)) (f64.const -1.0))
(assert_return (invoke "f64.convert_i32_u" (i32.const -1)) (f64.const 4294967295.0))
(assert_return (invoke "f64.convert_i64_s" (i64.const -1)) (f64.const -1.0))
(assert_return (invoke "f64.convert_i64_u" (i64.const -1)) (f64.const 18446744073709551616.0))
(assert_return (invoke "f32.convert_i64_u" (i64.const -1)) (f32.const 18446744073709551616.0))

;; ── the i32 extremes ────────────────────────────────────────────────────
(assert_return (invoke "f64.convert_i32_s" (i32.const 0x7fffffff)) (f64.const 2147483647.0))
(assert_return (invoke "f64.convert_i32_s" (i32.const 0x80000000)) (f64.const -2147483648.0))
(assert_return (invoke "f64.convert_i32_u" (i32.const 0x80000000)) (f64.const 2147483648.0))
;; f32 cannot hold 2147483647 — 24 mantissa bits — so it rounds to 2^31.
(assert_return (invoke "f32.convert_i32_s" (i32.const 0x7fffffff)) (f32.const 2147483648.0))
(assert_return (invoke "f32.convert_i32_s" (i32.const 0x80000000)) (f32.const -2147483648.0))

;; ── rounding is to nearest, TIES TO EVEN — not truncation ───────────────
;; 16777217 = 2^24+1 is the first integer f32 cannot represent. Ties-to-even
;; sends it DOWN to 16777216; a truncating implementation agrees here...
(assert_return (invoke "f32.convert_i32_s" (i32.const 16777217)) (f32.const 16777216.0))
;; ...but not here: 16777219 is nearer to 16777220 than to 16777216, so
;; rounding goes UP. Truncation would give 16777216 and fail.
(assert_return (invoke "f32.convert_i32_s" (i32.const 16777219)) (f32.const 16777220.0))
(assert_return (invoke "f32.convert_i32_s" (i32.const -16777219)) (f32.const -16777220.0))
;; The same boundary one binade higher.
(assert_return (invoke "f32.convert_i32_s" (i32.const 33554434)) (f32.const 33554432.0))
(assert_return (invoke "f32.convert_i32_s" (i32.const 33554436)) (f32.const 33554436.0))

;; ── f64 is exact for every i32, and lossy above 2^53 for i64 ────────────
(assert_return (invoke "f64.convert_i32_s" (i32.const 16777217)) (f64.const 16777217.0))
;; 2^53+1 is the first integer f64 cannot represent.
(assert_return (invoke "f64.convert_i64_s" (i64.const 9007199254740993)) (f64.const 9007199254740992.0))
(assert_return (invoke "f64.convert_i64_s" (i64.const 9007199254740995)) (f64.const 9007199254740996.0))
(assert_return (invoke "f64.convert_i64_s" (i64.const 0x7fffffffffffffff)) (f64.const 9223372036854775808.0))
(assert_return (invoke "f64.convert_i64_s" (i64.const 0x8000000000000000)) (f64.const -9223372036854775808.0))
(assert_return (invoke "f64.convert_i64_u" (i64.const 0x8000000000000000)) (f64.const 9223372036854775808.0))

;; ── i64 → f32 loses the most, and still rounds rather than truncates ────
(assert_return (invoke "f32.convert_i64_s" (i64.const 16777219)) (f32.const 16777220.0))
(assert_return (invoke "f32.convert_i64_s" (i64.const 0x7fffffffffffffff)) (f32.const 9223372036854775808.0))
(assert_return (invoke "f32.convert_i64_u" (i64.const 0x8000000000000000)) (f32.const 9223372036854775808.0))

;; ── zero keeps no sign: there is no negative zero integer ───────────────
(assert_return (invoke "f32.convert_i32_s" (i32.const 0)) (f32.const 0.0))
(assert_return (invoke "f64.convert_i64_u" (i64.const 0)) (f64.const 0.0))
