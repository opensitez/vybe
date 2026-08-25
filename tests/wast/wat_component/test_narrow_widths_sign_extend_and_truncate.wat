;; vybe-test: wast/wat_component/test_narrow_widths_sign_extend_and_truncate
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Loading (:2376) — the signed and unsigned widths are SEPARATE cases:
;;     case U8Type() : return load_int(cx, ptr, 1)
;;     case S8Type() : return load_int(cx, ptr, 1, signed = True)
;;   and §Flat Lifting (:3282).
;;
;; ▶▶ `s8` AND `u8` ARE NOT `I32` WITH A SMALLER RANGE. They used to refuse for
;; exactly that reason: an `I32` performs neither the narrowing nor the sign
;; extension, so a widened value crosses the ABI silently out of range.
;;
;; The discriminator is ONE BYTE PATTERN READ TWO WAYS. `0xFF` is:
;;
;;   s8  → -1     (sign-extended)
;;   u8  → 255    (zero-extended)
;;
;; so a single test value separates the two treatments, and separates BOTH from
;; the "just widen it to i32" answer, which would give 255 for `s8`.
;;
;; The core callee is handed the lowered args and returns `a + b`. With the
;; component signature `(param "a" s8) (param "b" u8)`:
;;
;;   -1 + 255 = 254   correct
;;   255 + 255 = 510  both treated as unsigned — the "widen it" bug
;;   -1 + -1  = -2    both treated as signed
;;
;; ⛔ The caller passes 0xFF in BOTH slots, so the two parameters are
;; indistinguishable at the core boundary and only the DECLARED type can tell
;; them apart. Passing different numbers would let a correct answer arise from
;; reading the slots in the wrong order.

(component
  (core module $m
    (func (export "add") (param i32 i32) (result i32)
      (i32.add (local.get 0) (local.get 1))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "add" (core func $a))

  (type $ft (func (param "a" s8) (param "b" u8) (result s32)))
  (canon lift (core func $a) (func $summed (type $ft)))
  (canon lower (func $summed) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32 i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 0xFF) (i32.const 0xFF))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 254))
