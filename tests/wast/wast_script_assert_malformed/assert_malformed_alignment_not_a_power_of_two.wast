;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_alignment_not_a_power_of_two
;; vybe-test-mode: run
;;
;; `align=N` states the access alignment in BYTES, and the binary format
;; encodes it as the EXPONENT: the memarg carries `k` for an alignment of
;; 2^k. A value that is not a positive power of two therefore has no
;; encoding at all, which is why the spec calls it MALFORMED text rather
;; than an invalid module (`align.wast`, `simd_align.wast`).
;;
;; Two things share the `align=` slot and must not be confused with it:
;; an alignment LARGER than the access is natural width — `i32.load8_u
;; align=2` — which encodes fine and is merely INVALID, and `offset=`,
;; which is an ordinary byte count with no power-of-two constraint.
;;
;; Our grammar takes `align=` as `"align=" ~ integer`, so every one of
;; these parsed and the whole `assert_malformed` half of `align.wast`
;; passed vacuously.

;; ── Zero: not a power of two, and no exponent encodes it ─────────────
(assert_malformed
  (module quote "(module (memory 0) (func (drop (i32.load8_s align=0 (i32.const 0)))))")
  "alignment")
(assert_malformed
  (module quote "(module (memory 0) (func (drop (i32.load align=0 (i32.const 0)))))")
  "alignment")
(assert_malformed
  (module quote "(module (memory 0) (func (i32.store align=0 (i32.const 0) (i32.const 1))))")
  "alignment")

;; ── An odd, non-power-of-two width ───────────────────────────────────
(assert_malformed
  (module quote "(module (memory 0) (func (drop (i32.load8_u align=7 (i32.const 0)))))")
  "alignment")
(assert_malformed
  (module quote "(module (memory 0) (func (drop (i64.load align=7 (i32.const 0)))))")
  "alignment")
(assert_malformed
  (module quote "(module (memory 0) (func (drop (f64.load align=3 (i32.const 0)))))")
  "alignment")

;; ── Negative: the grammar's `integer` accepts a sign ─────────────────
(assert_malformed
  (module quote "(module (memory 1) (func (drop (v128.load align=-1 (i32.const 0)))))")
  "alignment")

;; ── The PLAIN spelling reaches the same check ────────────────────────
(assert_malformed
  (module quote "(module (memory 0) (func (drop (i32.const 0) i32.load align=6)))")
  "alignment")

;; ── Controls: every legal power of two still parses ──────────────────
;; A check that rejects too much would take these with it.
(module (memory 0) (func (drop (i32.load8_s align=1 (i32.const 0)))))
(module (memory 0) (func (drop (i32.load16_u align=2 (i32.const 0)))))
(module (memory 0) (func (drop (i32.load align=4 (i32.const 0)))))
(module (memory 0) (func (drop (i64.load align=8 (i32.const 0)))))
(module (memory 1) (func (drop (v128.load align=16 (i32.const 0)))))
(module (memory 0) (func (i64.store align=8 (i32.const 0) (i64.const 1))))

;; An alignment wider than the access is NOT malformed — it parses, and is
;; rejected (if at all) by validation. Treating it as malformed here would
;; make `align.wast`'s `assert_invalid` half unreachable.
(module (memory 0) (func (drop (i32.load8_u align=2 (i32.const 0)))))
(module (memory 0) (func (drop (i32.load align=8 (i32.const 0)))))

;; `offset=` shares the slot and takes any byte count.
(module (memory 1) (func (drop (i32.load offset=3 (i32.const 0)))))
(module (memory 1) (func (drop (i32.load offset=7 align=4 (i32.const 0)))))

;; ── And the alignment is still only a hint at run time ───────────────
(module
  (memory 1)
  (func (export "roundtrip") (result i32)
    (i32.store align=1 (i32.const 1) (i32.const 0x11223344))
    (i32.load align=1 (i32.const 1))))
(assert_return (invoke "roundtrip") (i32.const 0x11223344))
