;; vybe-test: wast/wat_ops_simd/f32x4_lanes_nan_and_zero
;; origin: coverage gap — 14 f32x4 mnemonics occurred at most ONCE in the run corpus
;; vybe-test-mode: run
;;
;; f32x4 arithmetic at the points where a lane is not an ordinary number.
;;
;; The lane ops that move a scalar into or out of a vector — `splat`,
;; `replace_lane`, `extract_lane` — are BIT-PRESERVING. They are not
;; conversions, so a signalling NaN must arrive with its quiet bit still clear.
;; An implementation that carries the scalar through f64 on the way in (which
;; is invisible for every finite value, since f32→f64→f32 is exact) quiets it,
;; and ONLY an SNaN lane detects that.
;;
;; `min`/`max` are the other trap: they are not `a < b ? a : b`. They return
;; NaN if either lane is NaN, and they order -0.0 below +0.0 where `<` calls
;; the two equal. Lane-wise, that means a vector mixing signed zeros with NaN
;; distinguishes a correct implementation from a comparison-based one in a
;; single assertion.
;;
;; Bit patterns are read back through `i32x4` so signed zeros and NaN payloads
;; are visible; comparing as f32 would make -0.0 and +0.0 identical.

(module
  (func (export "splat") (param f32) (result v128) (f32x4.splat (local.get 0)))
  (func (export "ext0") (param v128) (result f32) (f32x4.extract_lane 0 (local.get 0)))
  (func (export "ext3") (param v128) (result f32) (f32x4.extract_lane 3 (local.get 0)))
  (func (export "replace2") (param v128 f32) (result v128) (f32x4.replace_lane 2 (local.get 0) (local.get 1)))
  (func (export "add") (param v128 v128) (result v128) (f32x4.add (local.get 0) (local.get 1)))
  (func (export "sub") (param v128 v128) (result v128) (f32x4.sub (local.get 0) (local.get 1)))
  (func (export "mul") (param v128 v128) (result v128) (f32x4.mul (local.get 0) (local.get 1)))
  (func (export "div") (param v128 v128) (result v128) (f32x4.div (local.get 0) (local.get 1)))
  (func (export "neg") (param v128) (result v128) (f32x4.neg (local.get 0)))
  (func (export "abs") (param v128) (result v128) (f32x4.abs (local.get 0)))
  (func (export "sqrt") (param v128) (result v128) (f32x4.sqrt (local.get 0)))
  (func (export "min") (param v128 v128) (result v128) (f32x4.min (local.get 0) (local.get 1)))
  (func (export "max") (param v128 v128) (result v128) (f32x4.max (local.get 0) (local.get 1)))
  (func (export "eq") (param v128 v128) (result v128) (f32x4.eq (local.get 0) (local.get 1)))
  (func (export "lt") (param v128 v128) (result v128) (f32x4.lt (local.get 0) (local.get 1)))
  ;; Reads the raw encoding of a splat — the only way to see a quieted NaN.
  (func (export "splat_bits") (param f32) (result v128) (f32x4.splat (local.get 0)))
)

;; ── splat and extract are bit-preserving ────────────────────────────────
(assert_return (invoke "splat" (f32.const 1.5))
               (v128.const f32x4 1.5 1.5 1.5 1.5))
(assert_return (invoke "ext0" (v128.const f32x4 1.0 2.0 3.0 4.0)) (f32.const 1.0))
;; The LAST lane, so an implementation that reads lane 0 regardless is caught.
(assert_return (invoke "ext3" (v128.const f32x4 1.0 2.0 3.0 4.0)) (f32.const 4.0))
(assert_return (invoke "replace2" (v128.const f32x4 0.0 0.0 0.0 0.0) (f32.const 7.0))
               (v128.const f32x4 0.0 0.0 7.0 0.0))
;; A SIGNALLING NaN through splat: the quiet bit must still be CLEAR. Read as
;; i32x4 because `f32x4.const nan:0x200000` would compare equal to any NaN.
(assert_return (invoke "splat_bits" (f32.const nan:0x200000))
               (v128.const i32x4 0x7fa00000 0x7fa00000 0x7fa00000 0x7fa00000))
(assert_return (invoke "splat_bits" (f32.const -nan))
               (v128.const i32x4 0xffc00000 0xffc00000 0xffc00000 0xffc00000))
;; ...and through replace_lane, which is the other scalar→lane path.
(assert_return (invoke "replace2" (v128.const i32x4 0 0 0 0) (f32.const nan:0x200000))
               (v128.const i32x4 0 0 0x7fa00000 0))
;; A signed zero survives too — invisible to any f32 comparison.
(assert_return (invoke "splat_bits" (f32.const -0.0))
               (v128.const i32x4 0x80000000 0x80000000 0x80000000 0x80000000))

;; ── ordinary arithmetic, lane by lane ───────────────────────────────────
(assert_return (invoke "add" (v128.const f32x4 1.0 2.0 3.0 4.0) (v128.const f32x4 10.0 20.0 30.0 40.0))
               (v128.const f32x4 11.0 22.0 33.0 44.0))
(assert_return (invoke "sub" (v128.const f32x4 1.0 2.0 3.0 4.0) (v128.const f32x4 10.0 20.0 30.0 40.0))
               (v128.const f32x4 -9.0 -18.0 -27.0 -36.0))
(assert_return (invoke "mul" (v128.const f32x4 1.5 2.0 -3.0 0.0) (v128.const f32x4 2.0 2.0 2.0 2.0))
               (v128.const f32x4 3.0 4.0 -6.0 0.0))

;; ── each lane rounds at 24 bits, independently ──────────────────────────
;; A lane computed in double precision keeps the 2^-24 term and fails.
(assert_return (invoke "add" (v128.const f32x4 1.0 1.0 1.0 1.0)
                             (v128.const f32x4 0x1p-24 0x1p-23 0x1.8p-24 0.0))
               (v128.const f32x4 1.0 0x1.000002p+0 0x1.000002p+0 1.0))
;; Overflow in one lane does not disturb its neighbours.
(assert_return (invoke "mul" (v128.const f32x4 0x1.fffffep+127 1.0 0x1p-149 1.0)
                             (v128.const f32x4 2.0 1.0 0.5 1.0))
               (v128.const f32x4 inf 1.0 0.0 1.0))

;; ── division: zeros, infinities and the NaN cases ───────────────────────
(assert_return (invoke "div" (v128.const f32x4 1.0 1.0 0.0 1.0)
                             (v128.const f32x4 0.0 -0.0 0.0 inf))
               (v128.const f32x4 inf -inf nan:canonical 0.0))

;; ── neg and abs are SIGN-BIT ops, including on NaN ──────────────────────
(assert_return (invoke "neg" (v128.const i32x4 0x7fc00000 0x00000000 0x3f800000 0x80000000))
               (v128.const i32x4 0xffc00000 0x80000000 0xbf800000 0x00000000))
(assert_return (invoke "abs" (v128.const i32x4 0xffc00000 0x80000000 0xbf800000 0x00000000))
               (v128.const i32x4 0x7fc00000 0x00000000 0x3f800000 0x00000000))

;; ── sqrt ────────────────────────────────────────────────────────────────
;; sqrt(-0.0) is -0.0, the one negative input that is not NaN.
(assert_return (invoke "sqrt" (v128.const f32x4 4.0 -1.0 inf 0.0))
               (v128.const f32x4 2.0 nan:canonical inf 0.0))

;; ── min / max are not `<` ───────────────────────────────────────────────
;; NaN in either operand wins, in either lane position.
(assert_return (invoke "min" (v128.const f32x4 nan 1.0 1.0 2.0)
                             (v128.const f32x4 1.0 nan 2.0 1.0))
               (v128.const f32x4 nan:canonical nan:canonical 1.0 1.0))
(assert_return (invoke "max" (v128.const f32x4 nan 1.0 1.0 2.0)
                             (v128.const f32x4 1.0 nan 2.0 1.0))
               (v128.const f32x4 nan:canonical nan:canonical 2.0 2.0))
;; -0.0 and +0.0 compare EQUAL, so a `<`-based min returns whichever came
;; first. Reading the bits is the only way to see that it picked correctly.
(assert_return (invoke "min" (v128.const f32x4 -0.0 0.0 -0.0 0.0)
                             (v128.const f32x4 0.0 -0.0 -0.0 0.0))
               (v128.const i32x4 0x80000000 0x80000000 0x80000000 0x00000000))
(assert_return (invoke "max" (v128.const f32x4 -0.0 0.0 -0.0 0.0)
                             (v128.const f32x4 0.0 -0.0 -0.0 0.0))
               (v128.const i32x4 0x00000000 0x00000000 0x80000000 0x00000000))

;; ── comparisons give all-ones / all-zeros masks, and NaN is never equal ─
(assert_return (invoke "eq" (v128.const f32x4 1.0 1.0 nan -0.0)
                            (v128.const f32x4 1.0 2.0 nan 0.0))
               (v128.const i32x4 -1 0 0 -1))
;; NaN is unordered: BOTH lt and eq are false against it.
(assert_return (invoke "lt" (v128.const f32x4 1.0 2.0 nan 1.0)
                            (v128.const f32x4 2.0 1.0 1.0 nan))
               (v128.const i32x4 -1 0 0 0))
