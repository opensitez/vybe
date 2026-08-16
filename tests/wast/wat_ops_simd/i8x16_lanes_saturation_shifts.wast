;; vybe-test: wast/wat_ops_simd/i8x16_lanes_saturation_shifts
;; origin: coverage gap — 22 i8x16 mnemonics occurred at most ONCE in the run corpus
;; vybe-test-mode: run
;;
;; i8x16 is where lane semantics are sharpest, because a byte overflows almost
;; immediately. Three properties, none reachable by a small-positive test:
;;
;;   * `extract_lane_s` vs `extract_lane_u` differ ONLY when the byte has its
;;     top bit set — 0x80 is -128 signed and 128 unsigned. Every lane value
;;     under 0x80 makes the two opcodes indistinguishable.
;;   * `add` WRAPS and `add_sat_s`/`add_sat_u` CLAMP. 127+1 is -128 wrapping
;;     but 127 saturating; that difference is the entire reason both exist.
;;   * shift counts are taken modulo the LANE width (8), not the vector width,
;;     so a shift by 8 is the identity and by 9 is a shift by 1.
;;
;; `bitmask` is included because it is the one op that reads the sign bit of
;; every lane at once, so a lane-ordering mistake shows up as a scrambled
;; integer rather than a wrong lane.

(module
  (func (export "splat") (param i32) (result v128) (i8x16.splat (local.get 0)))
  (func (export "ext_s0") (param v128) (result i32) (i8x16.extract_lane_s 0 (local.get 0)))
  (func (export "ext_u0") (param v128) (result i32) (i8x16.extract_lane_u 0 (local.get 0)))
  (func (export "ext_s15") (param v128) (result i32) (i8x16.extract_lane_s 15 (local.get 0)))
  (func (export "ext_u15") (param v128) (result i32) (i8x16.extract_lane_u 15 (local.get 0)))
  (func (export "replace3") (param v128 i32) (result v128) (i8x16.replace_lane 3 (local.get 0) (local.get 1)))
  (func (export "add") (param v128 v128) (result v128) (i8x16.add (local.get 0) (local.get 1)))
  (func (export "sub") (param v128 v128) (result v128) (i8x16.sub (local.get 0) (local.get 1)))
  (func (export "add_sat_s") (param v128 v128) (result v128) (i8x16.add_sat_s (local.get 0) (local.get 1)))
  (func (export "add_sat_u") (param v128 v128) (result v128) (i8x16.add_sat_u (local.get 0) (local.get 1)))
  (func (export "sub_sat_s") (param v128 v128) (result v128) (i8x16.sub_sat_s (local.get 0) (local.get 1)))
  (func (export "sub_sat_u") (param v128 v128) (result v128) (i8x16.sub_sat_u (local.get 0) (local.get 1)))
  (func (export "shl") (param v128 i32) (result v128) (i8x16.shl (local.get 0) (local.get 1)))
  (func (export "shr_s") (param v128 i32) (result v128) (i8x16.shr_s (local.get 0) (local.get 1)))
  (func (export "shr_u") (param v128 i32) (result v128) (i8x16.shr_u (local.get 0) (local.get 1)))
  (func (export "neg") (param v128) (result v128) (i8x16.neg (local.get 0)))
  (func (export "abs") (param v128) (result v128) (i8x16.abs (local.get 0)))
  (func (export "min_s") (param v128 v128) (result v128) (i8x16.min_s (local.get 0) (local.get 1)))
  (func (export "min_u") (param v128 v128) (result v128) (i8x16.min_u (local.get 0) (local.get 1)))
  (func (export "max_s") (param v128 v128) (result v128) (i8x16.max_s (local.get 0) (local.get 1)))
  (func (export "bitmask") (param v128) (result i32) (i8x16.bitmask (local.get 0)))
  (func (export "eq") (param v128 v128) (result v128) (i8x16.eq (local.get 0) (local.get 1)))
  (func (export "lt_s") (param v128 v128) (result v128) (i8x16.lt_s (local.get 0) (local.get 1)))
  (func (export "lt_u") (param v128 v128) (result v128) (i8x16.lt_u (local.get 0) (local.get 1)))
)

;; ── splat fills every lane; extract reads the one asked for ─────────────
(assert_return (invoke "splat" (i32.const 5))
               (v128.const i8x16 5 5 5 5 5 5 5 5 5 5 5 5 5 5 5 5))
;; splat TRUNCATES its i32 operand to 8 bits rather than saturating.
(assert_return (invoke "splat" (i32.const 0x1ff))
               (v128.const i8x16 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1))
(assert_return (invoke "replace3" (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0) (i32.const 7))
               (v128.const i8x16 0 0 0 7 0 0 0 0 0 0 0 0 0 0 0 0))

;; ── extract_lane_s vs _u: only a top-bit-set lane separates them ────────
(assert_return (invoke "ext_s0" (v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const -128))
(assert_return (invoke "ext_u0" (v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 128))
(assert_return (invoke "ext_s0" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const -1))
(assert_return (invoke "ext_u0" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 255))
;; A lane below 0x80 agrees between the two — the case that proves nothing.
(assert_return (invoke "ext_s0" (v128.const i8x16 0x7f 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 127))
(assert_return (invoke "ext_u0" (v128.const i8x16 0x7f 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 127))
;; The LAST lane, so an implementation that reads lane 0 regardless is caught.
(assert_return (invoke "ext_s15" (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0x80)) (i32.const -128))
(assert_return (invoke "ext_u15" (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0x80)) (i32.const 128))

;; ── add wraps; add_sat clamps ───────────────────────────────────────────
(assert_return (invoke "add" (v128.const i8x16 127 127 -128 -128 0 0 0 0 0 0 0 0 0 0 0 0)
                             (v128.const i8x16 1 -1 -1 1 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -128 126 127 -127 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "add_sat_s" (v128.const i8x16 127 127 -128 -128 0 0 0 0 0 0 0 0 0 0 0 0)
                                   (v128.const i8x16 1 -1 -1 1 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 127 126 -128 -127 0 0 0 0 0 0 0 0 0 0 0 0))
;; Unsigned saturation has a different ceiling (255) and floor (0).
(assert_return (invoke "add_sat_u" (v128.const i8x16 0xff 0xff 0x80 1 0 0 0 0 0 0 0 0 0 0 0 0)
                                   (v128.const i8x16 1 0xff 0x80 1 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 0xff 0xff 0xff 2 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "sub_sat_s" (v128.const i8x16 -128 -128 127 0 0 0 0 0 0 0 0 0 0 0 0 0)
                                   (v128.const i8x16 1 -1 -1 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -128 -127 127 0 0 0 0 0 0 0 0 0 0 0 0 0))
;; sub_sat_u floors at 0 — 0-1 is 0, not 255.
(assert_return (invoke "sub_sat_u" (v128.const i8x16 0 5 0xff 0 0 0 0 0 0 0 0 0 0 0 0 0)
                                   (v128.const i8x16 1 10 1 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 0 0 0xfe 0 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "sub" (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)
                             (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))

;; ── shift counts are modulo the LANE width (8) ──────────────────────────
(assert_return (invoke "shl" (v128.const i8x16 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1) (i32.const 1))
               (v128.const i8x16 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2))
;; Shift by 8 is the IDENTITY, not zero — a clamping implementation returns 0.
(assert_return (invoke "shl" (v128.const i8x16 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1) (i32.const 8))
               (v128.const i8x16 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1))
(assert_return (invoke "shl" (v128.const i8x16 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1) (i32.const 9))
               (v128.const i8x16 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2))
;; Bits shifted out of a lane are DISCARDED, not carried into the next lane.
(assert_return (invoke "shl" (v128.const i8x16 0x80 0 0x40 0 0 0 0 0 0 0 0 0 0 0 0 0) (i32.const 1))
               (v128.const i8x16 0 0 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0))
;; shr_s copies the sign bit in; shr_u brings zeros.
(assert_return (invoke "shr_s" (v128.const i8x16 -1 -128 0x40 0 0 0 0 0 0 0 0 0 0 0 0 0) (i32.const 1))
               (v128.const i8x16 -1 -64 0x20 0 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "shr_u" (v128.const i8x16 -1 -128 0x40 0 0 0 0 0 0 0 0 0 0 0 0 0) (i32.const 1))
               (v128.const i8x16 0x7f 0x40 0x20 0 0 0 0 0 0 0 0 0 0 0 0 0))

;; ── neg and abs at the asymmetric extreme ───────────────────────────────
;; -128 has no positive counterpart in an i8: both neg and abs return -128.
(assert_return (invoke "neg" (v128.const i8x16 -128 1 -1 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -128 -1 1 0 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "abs" (v128.const i8x16 -128 -1 1 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -128 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0))

;; ── signed vs unsigned min/max on the same lanes ────────────────────────
;; -1 is the LARGEST unsigned byte and the SMALLEST-but-one signed one.
(assert_return (invoke "min_s" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)
                               (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "min_u" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)
                               (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "max_s" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)
                               (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))

;; ── bitmask reads every lane's sign bit, in lane order ──────────────────
(assert_return (invoke "bitmask" (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 0))
(assert_return (invoke "bitmask" (v128.const i8x16 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1))
               (i32.const 0xffff))
;; Lane 0 is bit 0 — a reversed lane order gives 0x8000 here.
(assert_return (invoke "bitmask" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 1))
(assert_return (invoke "bitmask" (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 -1)) (i32.const 0x8000))
;; 0x7f has its sign bit CLEAR — "non-zero" is not the test.
(assert_return (invoke "bitmask" (v128.const i8x16 0x7f 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0)) (i32.const 2))

;; ── comparisons produce ALL-ONES / ALL-ZEROS lane masks ─────────────────
(assert_return (invoke "eq" (v128.const i8x16 1 2 3 4 0 0 0 0 0 0 0 0 0 0 0 0)
                            (v128.const i8x16 1 0 3 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -1 0 -1 0 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1 -1))
;; -1 < 1 signed, but 255 > 1 unsigned: same bytes, opposite masks.
(assert_return (invoke "lt_s" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)
                              (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "lt_u" (v128.const i8x16 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)
                              (v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
