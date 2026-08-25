;; vybe-test: wast/wat_component/test_narrow_width_truncates_the_slot
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Flat Lifting (:3282) and §Loading (:2377) — `u8` is `load_int(cx, ptr, 1)`.
;;
;; ▶▶ THIS FILE REPLACES A REFUSAL. It used to be
;; `test_component_type_refuses_an_uncarryable_width`, pinning the compile
;; error `u8` produced when `component::ValType` had only `I32`/`I64`. Its own
;; header said: *"This pins a REFUSAL, not a feature — when
;; `component::ValType` grows the narrow widths … this test should be REPLACED
;; by one that round-trips a `u8`, not deleted."* The widths landed, so here is
;; that round trip, asserting the exact claim the refusal used to make.
;;
;; THE CLAIM WAS TRUNCATION: *"lifting a u8 requires truncation to eight bits,
;; which an `I32` does not perform, so the value arriving on the other side
;; would be silently out of range."*
;;
;; So the caller passes **0x1FF = 511**, which does not fit in eight bits. A
;; `u8` must narrow it to `0xFF = 255`; an `I32` stand-in passes 511 straight
;; through. The callee returns its argument unchanged, so:
;;
;;   255  correct — the high bit of the slot is NOT part of the value
;;   511  the value was never narrowed
;;   -1   narrowed but sign-extended, i.e. `s8` semantics on a `u8`
;;
;; ⛔ Distinct from `test_narrow_widths_sign_extend_and_truncate`, which proves
;; SIGN EXTENSION with a value already inside eight bits. This one proves the
;; slot's spare bits are discarded. Neither implies the other: a stand-in that
;; masked but did not sign-extend would pass this file and fail that one.

(component
  (core module $m
    (func (export "id") (param i32) (result i32) (local.get 0)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "id" (core func $d))

  (type $ft (func (param "a" u8) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))
  (canon lower (func $lifted) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 0x1FF))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 255))
