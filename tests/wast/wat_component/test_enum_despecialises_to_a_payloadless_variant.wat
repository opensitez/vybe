;; vybe-test: wast/wat_component/test_enum_despecialises_to_a_payloadless_variant
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Despecialization (:2175):
;;     case EnumType(labels) : return VariantType([ CaseType(l, None) for l in labels ])
;;
;; ▶▶ `enum` is a variant whose every case is payload-less, so it flattens to
;; exactly ONE i32 — the discriminant — and nothing else.
;;
;; The core callee takes one i32 and the lowered core function does too, so an
;; enum that despecialised into anything with a payload would flatten to two
;; values and fail on arity rather than on the answer.
;;
;; `blue` is case **2 of 3**, and the callee multiplies by 10, so 20 can only
;; come from the discriminant surviving the round trip intact: case 0 gives 0
;; and case 1 gives 10. The companion file
;; `test_an_out_of_range_enum_case_traps` pins the other half of the claim —
;; that the case COUNT reached the runtime, not just the shape.

(component
  (core module $m
    (func (export "scale") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 10))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "scale" (core func $s))

  (type $ft (func (param "e" (enum "red" "green" "blue")) (result u32)))
  (canon lift (core func $s) (func $scaled (type $ft)))
  (canon lower (func $scaled) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 2))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 20))
