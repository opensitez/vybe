;; vybe-test: wast/wat_simd_full_vector/load_store_lane_splat_and_widening_loads
;; vybe-test-mode: run
;;
;; The v128 memory family: `v128.load`/`store`, the `_splat` loads, the widening
;; `load8x8`/`load16x4`/`load32x2` in both signednesses, `load32_zero`/
;; `load64_zero`, and the eight per-lane `load*_lane` / `store*_lane` forms.
;; Twenty of these occurred once each in the corpus, on data whose bytes were
;; small and positive — where signed and unsigned widening agree, so nothing
;; separated `load8x8_s` from `load8x8_u`, and every splat looked like every
;; other splat.
;;
;; The data segment is deliberately two halves: bytes 0–7 are 1…8 (small,
;; positive, and ASCENDING so a wrong byte order shows), bytes 8–15 are
;; 0xF0…0xF7 (every one of them negative as a signed lane). Every widening load
;; reads the SECOND half, so `_s` and `_u` cannot agree.
;;
;; The `_lane` forms are loaded into an all-zero vector and stored into a zeroed
;; region, so a write to the wrong lane leaves a zero where the value belongs
;; AND a value where a zero belongs — it cannot pass by touching the right byte
;; for the wrong reason.
;;
;; Spec-format so `wasmtime wast` arbitrates every expectation.

(module
  (memory (export "mem") 1)
  (data (i32.const 0) "\01\02\03\04\05\06\07\08\f0\f1\f2\f3\f4\f5\f6\f7")

  ;; ── Whole-vector load / store ────────────────────────────────────────
  (func (export "load") (result v128) (v128.load (i32.const 0)))
  (func (export "load_offset") (result v128) (v128.load offset=8 (i32.const 0)))
  (func (export "store_then_load") (result v128)
    (v128.store (i32.const 64) (v128.const i32x4 -1 2 -3 4))
    (v128.load (i32.const 64)))

  ;; ── Splats: one element broadcast to every lane ──────────────────────
  (func (export "load8_splat") (result v128) (v128.load8_splat (i32.const 8)))
  (func (export "load16_splat") (result v128) (v128.load16_splat (i32.const 8)))
  (func (export "load32_splat") (result v128) (v128.load32_splat (i32.const 8)))
  (func (export "load64_splat") (result v128) (v128.load64_splat (i32.const 0)))

  ;; ── Widening loads: half a vector of memory, extended to full lanes ──
  (func (export "load8x8_s") (result v128) (v128.load8x8_s (i32.const 8)))
  (func (export "load8x8_u") (result v128) (v128.load8x8_u (i32.const 8)))
  (func (export "load16x4_s") (result v128) (v128.load16x4_s (i32.const 8)))
  (func (export "load16x4_u") (result v128) (v128.load16x4_u (i32.const 8)))
  (func (export "load32x2_s") (result v128) (v128.load32x2_s (i32.const 8)))
  (func (export "load32x2_u") (result v128) (v128.load32x2_u (i32.const 8)))

  ;; ── Zero-extending scalar loads ──────────────────────────────────────
  (func (export "load32_zero") (result v128) (v128.load32_zero (i32.const 8)))
  (func (export "load64_zero") (result v128) (v128.load64_zero (i32.const 0)))

  ;; ── load*_lane: replace ONE lane, leave the rest ─────────────────────
  (func (export "load8_lane") (result v128)
    (v128.load8_lane 15 (i32.const 8) (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)))
  (func (export "load16_lane") (result v128)
    (v128.load16_lane 3 (i32.const 8) (v128.const i16x8 0 0 0 0 0 0 0 0)))
  (func (export "load32_lane") (result v128)
    (v128.load32_lane 2 (i32.const 8) (v128.const i32x4 0 0 0 0)))
  (func (export "load64_lane") (result v128)
    (v128.load64_lane 1 (i32.const 0) (v128.const i64x2 0 0)))

  ;; A lane load must PRESERVE the lanes it does not touch.
  (func (export "load32_lane_keeps_others") (result v128)
    (v128.load32_lane 1 (i32.const 8) (v128.const i32x4 11 22 33 44)))

  ;; ── store*_lane: write ONE lane into a zeroed region, read it back ───
  ;; Reading the surrounding bytes too, so a lane written one byte off is
  ;; visible rather than silently absorbed.
  (func (export "store8_lane") (result i32)
    (v128.store8_lane 5 (i32.const 96) (v128.const i8x16 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15))
    (i32.load offset=96 (i32.const 0)))
  (func (export "store16_lane") (result i32)
    (v128.store16_lane 2 (i32.const 112) (v128.const i16x8 100 200 300 400 500 600 700 800))
    (i32.load offset=112 (i32.const 0)))
  (func (export "store32_lane") (result i32)
    (v128.store32_lane 1 (i32.const 128) (v128.const i32x4 -1 -2 -3 -4))
    (i32.load offset=128 (i32.const 0)))
  (func (export "store64_lane") (result i64)
    (v128.store64_lane 1 (i32.const 144) (v128.const i64x2 7 -9))
    (i64.load offset=144 (i32.const 0)))

  ;; ── Out-of-bounds traps, at the exact byte where each width crosses ──
  (func (export "load_oob") (result v128) (v128.load (i32.const 65521)))
  (func (export "load64_zero_oob") (result v128) (v128.load64_zero (i32.const 65529)))
  (func (export "load8_lane_oob") (result v128)
    (v128.load8_lane 0 (i32.const 65536) (v128.const i64x2 0 0)))
  (func (export "store32_lane_oob")
    (v128.store32_lane 0 (i32.const 65533) (v128.const i32x4 1 2 3 4)))
)

;; ── whole vector ─────────────────────────────────────────────────────
(assert_return (invoke "load")
  (v128.const i8x16 1 2 3 4 5 6 7 8 -16 -15 -14 -13 -12 -11 -10 -9))
;; Only 8 bytes of data live past offset 8; the rest of the page is zero.
(assert_return (invoke "load_offset")
  (v128.const i8x16 -16 -15 -14 -13 -12 -11 -10 -9 0 0 0 0 0 0 0 0))
(assert_return (invoke "store_then_load") (v128.const i32x4 -1 2 -3 4))

;; ── splats ───────────────────────────────────────────────────────────
(assert_return (invoke "load8_splat")
  (v128.const i8x16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16 -16))
(assert_return (invoke "load16_splat")
  (v128.const i16x8 -3600 -3600 -3600 -3600 -3600 -3600 -3600 -3600))
(assert_return (invoke "load32_splat")
  (v128.const i32x4 -202182160 -202182160 -202182160 -202182160))
(assert_return (invoke "load64_splat")
  (v128.const i64x2 578437695752307201 578437695752307201))

;; ── widening loads: `_s` and `_u` differ on every lane ───────────────
(assert_return (invoke "load8x8_s") (v128.const i16x8 -16 -15 -14 -13 -12 -11 -10 -9))
(assert_return (invoke "load8x8_u") (v128.const i16x8 240 241 242 243 244 245 246 247))
(assert_return (invoke "load16x4_s") (v128.const i32x4 -3600 -3086 -2572 -2058))
(assert_return (invoke "load16x4_u") (v128.const i32x4 61936 62450 62964 63478))
(assert_return (invoke "load32x2_s") (v128.const i64x2 -202182160 -134810124))
(assert_return (invoke "load32x2_u") (v128.const i64x2 4092785136 4160157172))

;; ── zero-extending scalar loads ──────────────────────────────────────
(assert_return (invoke "load32_zero") (v128.const i32x4 -202182160 0 0 0))
(assert_return (invoke "load64_zero") (v128.const i64x2 578437695752307201 0))

;; ── load*_lane ───────────────────────────────────────────────────────
(assert_return (invoke "load8_lane")
  (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 -16))
(assert_return (invoke "load16_lane") (v128.const i16x8 0 0 0 -3600 0 0 0 0))
(assert_return (invoke "load32_lane") (v128.const i32x4 0 0 -202182160 0))
(assert_return (invoke "load64_lane") (v128.const i64x2 0 578437695752307201))
(assert_return (invoke "load32_lane_keeps_others")
  (v128.const i32x4 11 -202182160 33 44))

;; ── store*_lane ──────────────────────────────────────────────────────
;; Each reads back the whole 4/8-byte window, so the neighbouring zeros are
;; part of the assertion.
(assert_return (invoke "store8_lane") (i32.const 5))
(assert_return (invoke "store16_lane") (i32.const 300))
(assert_return (invoke "store32_lane") (i32.const -2))
(assert_return (invoke "store64_lane") (i64.const -9))

;; ── traps ────────────────────────────────────────────────────────────
(assert_trap (invoke "load_oob") "out of bounds memory access")
(assert_trap (invoke "load64_zero_oob") "out of bounds memory access")
(assert_trap (invoke "load8_lane_oob") "out of bounds memory access")
(assert_trap (invoke "store32_lane_oob") "out of bounds memory access")
