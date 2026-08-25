;; vybe-test: wast/wat_component/test_an_out_of_range_enum_case_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Despecialization (:2175) and §Flat Lifting (:3299).
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   canon lower: lifting arguments: canonical ABI: variant discriminant 3
;;   names no case; the type has 3 (0..2)
;;
;; ▶▶ THIS IS THE HALF `test_enum_despecialises_to_a_payloadless_variant`
;; CANNOT PROVE. That file shows an enum flattens to one i32 and the value
;; survives — but a variant with the WRONG NUMBER of cases would pass it too,
;; because case 2 is in range for any enum of three or more. Only an
;; out-of-range index shows the case COUNT reached the runtime, so the two
;; files together pin `enum "red" "green" "blue"` to exactly three cases.
;;
;; ⛔ THE MESSAGE ITSELF WAS THE OTHER DEFECT. It used to read "no store/load
;; implemented for variant discriminant out of range" — a
;; `CanonError::Unsupported`, whose Display says the FEATURE is missing. The
;; feature works; the PROGRAM is invalid, and a reader told otherwise goes
;; looking for unwritten code. It is now its own variant carrying both numbers,
;; and `option`/`result`/`enum` all despecialise through it.

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
    (func (export "_start")
      (drop (call $l (i32.const 3)))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)
