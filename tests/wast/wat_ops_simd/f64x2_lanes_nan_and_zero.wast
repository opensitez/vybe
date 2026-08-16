;; vybe-test: wast/wat_ops_simd/f64x2_lanes_nan_and_zero
;; origin: coverage gap — 26 f64x2 mnemonics occurred at most ONCE in the run corpus
;; vybe-test-mode: run
;;
;; f64x2 is the two-lane float shape. Everything said about f32x4 applies, with
;; one shape-specific risk on top: with only TWO lanes, an implementation that
;; broadcasts lane 0 or reads the wrong half is right half the time, so every
;; assertion below gives the two lanes DIFFERENT values.
;;
;; `min`/`max` remain the interesting ops — NaN wins over any operand, and
;; -0.0 orders below +0.0 where `<` calls them equal. Signed zeros are read
;; back as i64x2 because comparing as f64 cannot see the sign.

(module
  (func (export "splat") (param f64) (result v128) (f64x2.splat (local.get 0)))
  (func (export "ext0") (param v128) (result f64) (f64x2.extract_lane 0 (local.get 0)))
  (func (export "ext1") (param v128) (result f64) (f64x2.extract_lane 1 (local.get 0)))
  (func (export "replace1") (param v128 f64) (result v128) (f64x2.replace_lane 1 (local.get 0) (local.get 1)))
  (func (export "add") (param v128 v128) (result v128) (f64x2.add (local.get 0) (local.get 1)))
  (func (export "sub") (param v128 v128) (result v128) (f64x2.sub (local.get 0) (local.get 1)))
  (func (export "mul") (param v128 v128) (result v128) (f64x2.mul (local.get 0) (local.get 1)))
  (func (export "div") (param v128 v128) (result v128) (f64x2.div (local.get 0) (local.get 1)))
  (func (export "neg") (param v128) (result v128) (f64x2.neg (local.get 0)))
  (func (export "abs") (param v128) (result v128) (f64x2.abs (local.get 0)))
  (func (export "sqrt") (param v128) (result v128) (f64x2.sqrt (local.get 0)))
  (func (export "min") (param v128 v128) (result v128) (f64x2.min (local.get 0) (local.get 1)))
  (func (export "max") (param v128 v128) (result v128) (f64x2.max (local.get 0) (local.get 1)))
  (func (export "eq") (param v128 v128) (result v128) (f64x2.eq (local.get 0) (local.get 1)))
  (func (export "ne") (param v128 v128) (result v128) (f64x2.ne (local.get 0) (local.get 1)))
  (func (export "lt") (param v128 v128) (result v128) (f64x2.lt (local.get 0) (local.get 1)))
  (func (export "le") (param v128 v128) (result v128) (f64x2.le (local.get 0) (local.get 1)))
)

;; ── splat / extract / replace, with DISTINCT lanes throughout ───────────
(assert_return (invoke "splat" (f64.const 1.5)) (v128.const f64x2 1.5 1.5))
(assert_return (invoke "ext0" (v128.const f64x2 1.0 2.0)) (f64.const 1.0))
(assert_return (invoke "ext1" (v128.const f64x2 1.0 2.0)) (f64.const 2.0))
(assert_return (invoke "replace1" (v128.const f64x2 1.0 2.0) (f64.const 9.0))
               (v128.const f64x2 1.0 9.0))
;; Bit-preserving: a signalling NaN keeps its clear quiet bit through splat.
(assert_return (invoke "splat" (f64.const nan:0x4000000000000))
               (v128.const i64x2 0x7ff4000000000000 0x7ff4000000000000))
(assert_return (invoke "splat" (f64.const -0.0))
               (v128.const i64x2 0x8000000000000000 0x8000000000000000))
(assert_return (invoke "replace1" (v128.const i64x2 0 0) (f64.const -nan))
               (v128.const i64x2 0 0xfff8000000000000))

;; ── arithmetic, lane by lane ────────────────────────────────────────────
(assert_return (invoke "add" (v128.const f64x2 1.0 2.0) (v128.const f64x2 10.0 20.0))
               (v128.const f64x2 11.0 22.0))
(assert_return (invoke "sub" (v128.const f64x2 1.0 2.0) (v128.const f64x2 10.0 20.0))
               (v128.const f64x2 -9.0 -18.0))
(assert_return (invoke "mul" (v128.const f64x2 1.5 -3.0) (v128.const f64x2 2.0 2.0))
               (v128.const f64x2 3.0 -6.0))
;; f64 keeps 53 bits, so 1 + 2^-53 rounds away but 1 + 2^-52 does not — the
;; f32x4 boundary (2^-24) would be wrong here, catching a mis-sized lane.
(assert_return (invoke "add" (v128.const f64x2 1.0 1.0) (v128.const f64x2 0x1p-53 0x1p-52))
               (v128.const f64x2 1.0 0x1.0000000000001p+0))
;; Overflow in one lane leaves the other alone.
(assert_return (invoke "mul" (v128.const f64x2 0x1.fffffffffffffp+1023 1.0) (v128.const f64x2 2.0 1.0))
               (v128.const f64x2 inf 1.0))

;; ── division: zeros, infinities, NaN ────────────────────────────────────
(assert_return (invoke "div" (v128.const f64x2 1.0 1.0) (v128.const f64x2 0.0 -0.0))
               (v128.const f64x2 inf -inf))
(assert_return (invoke "div" (v128.const f64x2 0.0 inf) (v128.const f64x2 0.0 inf))
               (v128.const f64x2 nan:canonical nan:canonical))

;; ── neg / abs are SIGN-BIT ops, including on NaN ────────────────────────
(assert_return (invoke "neg" (v128.const i64x2 0x7ff8000000000000 0x0000000000000000))
               (v128.const i64x2 0xfff8000000000000 0x8000000000000000))
(assert_return (invoke "abs" (v128.const i64x2 0xfff8000000000000 0x8000000000000000))
               (v128.const i64x2 0x7ff8000000000000 0x0000000000000000))

;; ── sqrt: -0.0 is the one negative input that is not NaN ────────────────
(assert_return (invoke "sqrt" (v128.const f64x2 4.0 -1.0))
               (v128.const f64x2 2.0 nan:canonical))
(assert_return (invoke "sqrt" (v128.const f64x2 -0.0 inf))
               (v128.const i64x2 0x8000000000000000 0x7ff0000000000000))

;; ── min / max are not `<` ───────────────────────────────────────────────
;; NaN wins from EITHER side — the two lanes put it in opposite positions.
(assert_return (invoke "min" (v128.const f64x2 nan 1.0) (v128.const f64x2 1.0 nan))
               (v128.const f64x2 nan:canonical nan:canonical))
(assert_return (invoke "max" (v128.const f64x2 nan 1.0) (v128.const f64x2 1.0 nan))
               (v128.const f64x2 nan:canonical nan:canonical))
(assert_return (invoke "min" (v128.const f64x2 1.0 2.0) (v128.const f64x2 2.0 1.0))
               (v128.const f64x2 1.0 1.0))
(assert_return (invoke "max" (v128.const f64x2 1.0 2.0) (v128.const f64x2 2.0 1.0))
               (v128.const f64x2 2.0 2.0))
;; Signed zeros compare EQUAL, so only the bits show which operand was picked.
(assert_return (invoke "min" (v128.const f64x2 -0.0 0.0) (v128.const f64x2 0.0 -0.0))
               (v128.const i64x2 0x8000000000000000 0x8000000000000000))
(assert_return (invoke "max" (v128.const f64x2 -0.0 0.0) (v128.const f64x2 0.0 -0.0))
               (v128.const i64x2 0x0000000000000000 0x0000000000000000))

;; ── comparisons: all-ones / all-zeros masks, NaN unordered ──────────────
(assert_return (invoke "eq" (v128.const f64x2 1.0 -0.0) (v128.const f64x2 1.0 0.0))
               (v128.const i64x2 -1 -1))
;; NaN is equal to nothing, INCLUDING itself, and `ne` is correspondingly true.
(assert_return (invoke "eq" (v128.const f64x2 nan 1.0) (v128.const f64x2 nan 2.0))
               (v128.const i64x2 0 0))
(assert_return (invoke "ne" (v128.const f64x2 nan 1.0) (v128.const f64x2 nan 1.0))
               (v128.const i64x2 -1 0))
;; Unordered means BOTH lt and le are false, not just lt.
(assert_return (invoke "lt" (v128.const f64x2 nan 1.0) (v128.const f64x2 1.0 2.0))
               (v128.const i64x2 0 -1))
(assert_return (invoke "le" (v128.const f64x2 nan 1.0) (v128.const f64x2 nan 1.0))
               (v128.const i64x2 0 -1))
