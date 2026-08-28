;; vybe-test: wast/wat_simd_load_store/lane_ops_fold_the_offset_into_the_address
;; vybe-test-mode: run
;;
;; `offset=N` is a MEMARG: it is added to the ADDRESS operand, never to an
;; instruction's other immediates. Every load and store in the walker folds a
;; constant `offset=` into its address operand rather than carrying a runtime
;; memarg, and that fold wrote into operand slot 0.
;;
;; For `i32.load offset=5` slot 0 IS the address. For the eight SIMD lane ops
;; it is not: `v128.load8_lane` and friends carry a LANE INDEX immediate, and
;; the walker keeps immediates ahead of stack operands, so their operand list
;; is `[lane, addr, vec]`. The offset was landing on the LANE — `offset=3 1`
;; became lane 4 at address 0 instead of lane 1 at address 3.
;;
;; It hid because it needs BOTH forms to show: the plain `v128.load8_lane 1
;; (i32.const 3) …` spelling has no memarg to misplace and was always right, so
;; every lane op appeared in the corpus and appeared to work. All eight of the
;; spec's own lane files failed, and only on their `_offset_` exports.
;;
;; Each case below is written so the buggy lowering produces a DIFFERENT vector
;; rather than a coincidentally equal one: the lane the offset would shift to is
;; distinct from the lane named, and the byte at the un-offset address is
;; distinct from the byte at the offset one.

(module
  (memory 1)
  ;; Bytes 0..15 are their own index, so "which address was read" is legible
  ;; straight off the result. 16.. stays zero as a scratch area for the stores.
  (data (i32.const 0) "\00\01\02\03\04\05\06\07\08\09\0a\0b\0c\0d\0e\0f")

  ;; ── loads ────────────────────────────────────────────────────────────
  ;; lane 1 ← byte 3. Under the bug: lane 4 ← byte 0.
  (func (export "l8_off") (param $x v128) (result v128)
    (v128.load8_lane offset=3 1 (i32.const 0) (local.get $x)))
  ;; The same read with no memarg — the control that was already passing.
  (func (export "l8_plain") (param $x v128) (result v128)
    (v128.load8_lane 1 (i32.const 3) (local.get $x)))

  ;; lane 1 of i16x8 ← the u16 at 2 (= 0x0302). Under the bug: lane 3 ← u16 at 0.
  (func (export "l16_off") (param $x v128) (result v128)
    (v128.load16_lane offset=2 1 (i32.const 0) (local.get $x)))

  ;; lane 1 of i32x4 ← the u32 at 4 (= 0x07060504). Under the bug: lane 5 —
  ;; which is out of range for i32x4 entirely — at address 0.
  (func (export "l32_off") (param $x v128) (result v128)
    (v128.load32_lane offset=4 1 (i32.const 0) (local.get $x)))

  ;; lane 1 of i64x2 ← the u64 at 8. Under the bug: lane 9 at address 0.
  (func (export "l64_off") (param $x v128) (result v128)
    (v128.load64_lane offset=8 1 (i32.const 0) (local.get $x)))

  ;; ── stores ───────────────────────────────────────────────────────────
  ;; lane 3 → address 16+5 = 21. Under the bug: lane 8 → address 16.
  ;; Lanes 3 and 8 hold different values, and both target addresses are read
  ;; back, so either half of the mistake is visible on its own.
  (func (export "s8_off") (param $x v128)
    (v128.store8_lane offset=5 3 (i32.const 16) (local.get $x)))

  ;; lane 1 of i16x8 → address 32+2 = 34. Under the bug: lane 3 → address 32.
  (func (export "s16_off") (param $x v128)
    (v128.store16_lane offset=2 1 (i32.const 32) (local.get $x)))

  ;; lane 1 of i32x4 → address 48+4 = 52. Under the bug: lane 5 → address 48.
  (func (export "s32_off") (param $x v128)
    (v128.store32_lane offset=4 1 (i32.const 48) (local.get $x)))

  ;; lane 1 of i64x2 → address 64+8 = 72. Under the bug: lane 9 → address 64.
  (func (export "s64_off") (param $x v128)
    (v128.store64_lane offset=8 1 (i32.const 64) (local.get $x)))

  (func (export "b") (param $at i32) (result i32) (i32.load8_u (local.get $at)))

  ;; ── the control the fix must not break ───────────────────────────────
  ;; A v128 memory op with NO lane immediate still folds into slot 0. If the
  ;; lane discrimination were widened to the whole v128 family, this reads
  ;; address 0 instead of address 4.
  (func (export "v128_load_off") (result v128)
    (v128.load offset=4 (i32.const 0)))
  ;; …and so does a core load, which shares the same helper.
  (func (export "i32_load_off") (result i32)
    (i32.load offset=4 (i32.const 0)))
  ;; `load32_splat` contains neither "_lane" nor a lane immediate.
  (func (export "splat_off") (result v128)
    (v128.load32_splat offset=4 (i32.const 0)))
)

;; ── loads ──────────────────────────────────────────────────────────────
(assert_return (invoke "l8_off" (v128.const i8x16 0x77 0x77 0x77 0x77 0x77 0x77 0x77 0x77
                                               0x77 0x77 0x77 0x77 0x77 0x77 0x77 0x77))
                                (v128.const i8x16 0x77 3 0x77 0x77 0x77 0x77 0x77 0x77
                                               0x77 0x77 0x77 0x77 0x77 0x77 0x77 0x77))
(assert_return (invoke "l8_plain" (v128.const i8x16 0x77 0x77 0x77 0x77 0x77 0x77 0x77 0x77
                                                 0x77 0x77 0x77 0x77 0x77 0x77 0x77 0x77))
                                  (v128.const i8x16 0x77 3 0x77 0x77 0x77 0x77 0x77 0x77
                                                 0x77 0x77 0x77 0x77 0x77 0x77 0x77 0x77))

(assert_return (invoke "l16_off" (v128.const i16x8 0x7777 0x7777 0x7777 0x7777
                                                0x7777 0x7777 0x7777 0x7777))
                                 (v128.const i16x8 0x7777 0x0302 0x7777 0x7777
                                                0x7777 0x7777 0x7777 0x7777))

(assert_return (invoke "l32_off" (v128.const i32x4 0x77777777 0x77777777 0x77777777 0x77777777))
                                 (v128.const i32x4 0x77777777 0x07060504 0x77777777 0x77777777))

(assert_return (invoke "l64_off" (v128.const i64x2 0x7777777777777777 0x7777777777777777))
                                 (v128.const i64x2 0x7777777777777777 0x0f0e0d0c0b0a0908))

;; ── stores ─────────────────────────────────────────────────────────────
;; Lane 3 is 0xaa, lane 8 is 0xbb: whichever lane the store reads is named by
;; the byte that lands, and whichever address it writes is named by which read
;; back non-zero.
(invoke "s8_off" (v128.const i8x16 0 0 0 0xaa 0 0 0 0 0xbb 0 0 0 0 0 0 0))
(assert_return (invoke "b" (i32.const 21)) (i32.const 0xaa))
(assert_return (invoke "b" (i32.const 16)) (i32.const 0))

(invoke "s16_off" (v128.const i16x8 0 0xaabb 0 0xccdd 0 0 0 0))
(assert_return (invoke "b" (i32.const 34)) (i32.const 0xbb))
(assert_return (invoke "b" (i32.const 35)) (i32.const 0xaa))
(assert_return (invoke "b" (i32.const 32)) (i32.const 0))

(invoke "s32_off" (v128.const i32x4 0 0xaabbccdd 0 0))
(assert_return (invoke "b" (i32.const 52)) (i32.const 0xdd))
(assert_return (invoke "b" (i32.const 55)) (i32.const 0xaa))
(assert_return (invoke "b" (i32.const 48)) (i32.const 0))

(invoke "s64_off" (v128.const i64x2 0 0x1122334455667788))
(assert_return (invoke "b" (i32.const 72)) (i32.const 0x88))
(assert_return (invoke "b" (i32.const 79)) (i32.const 0x11))
(assert_return (invoke "b" (i32.const 64)) (i32.const 0))

;; ── controls ───────────────────────────────────────────────────────────
(assert_return (invoke "v128_load_off")
               (v128.const i8x16 4 5 6 7 8 9 10 11 12 13 14 15 0 0 0 0))
(assert_return (invoke "i32_load_off") (i32.const 0x07060504))
(assert_return (invoke "splat_off")
               (v128.const i32x4 0x07060504 0x07060504 0x07060504 0x07060504))
