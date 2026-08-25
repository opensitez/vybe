;; vybe-test: wast/wat_component/test_a_lone_surrogate_is_not_a_char
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md:2456
;;
;;   def convert_i32_to_char(cx, i):
;;     assert(i >= 0)
;;     trap_if(i >= 0x110000)
;;     trap_if(0xD800 <= i <= 0xDFFF)
;;     return chr(i)
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   canonical ABI: 0xd800 is not a Unicode scalar value — a `char` must be
;;   below 0x110000 and outside the surrogate range 0xD800..=0xDFFF
;;
;; ▶▶ THIS TRAP *IS* THE TYPE. `char` and `u32` are both four bytes in memory
;; and both flatten to one core `i32` — `flatten_type` cannot tell them apart,
;; `alignment` cannot, `elem_size` cannot. The ONLY thing that distinguishes
;; them is this validity check, which is why `char` refused rather than being
;; mapped onto `I32`: an `I32` stand-in carries a lone surrogate across the ABI
;; as though it were a character, and every layer downstream believes it.
;;
;; 0xD800 is the FIRST surrogate. It is a legal `u32` and a legal `i32` slot
;; value, so nothing but the range check can reject it — which is exactly the
;; property under test. A companion positive case is
;; `test_narrow_widths_sign_extend_and_truncate`; this file pins the refusal.

(component
  (core module $m
    (func (export "id") (param i32) (result i32) (local.get 0)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "id" (core func $d))

  (type $ft (func (param "c" char) (result s32)))
  (canon lift (core func $d) (func $lifted (type $ft)))
  (canon lower (func $lifted) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "_start")
      (drop (call $l (i32.const 0xD800)))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)
