;; vybe-test: wast/wat_component/test_an_import_occupies_a_component_func_index
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2601
;;   and :2611 — import ::= (import <externnamelit> <attribute>* bind-id(<externtype>))
;;
;; ▶▶ AN IMPORT TAKES A SLOT IT DOES NOT FILL. Imports and exports are on the
;; same footing in the spec — both "append a new element to the index space of
;; the imported/exported `sort`" — but an import supplies no definition. So the
;; slot must EXIST and must be EMPTY.
;;
;;   funcidx 0 → the IMPORT     (empty — nothing here defines it)
;;   funcidx 1 → the `canon lift`
;;
;; and `canon lower` names funcidx **1 POSITIONALLY**. That is the whole test:
;; if an imported function were skipped rather than slotted, the lift would sit
;; at 0, funcidx 1 would not exist, and this refuses with `component func 1 is
;; not defined (have 1)`. Skipping would renumber every function declared after
;; an import — correct-looking right up until two of them share a signature,
;; which is the `GLOBAL_GET` shape one more time.
;;
;; ⛔ THE IMPORT IS DECLARED AND NEVER CALLED, ON PURPOSE. A component may
;; legitimately import something it does not use, so the refusal for an
;; imported callee belongs at the CALL, not at the declaration — refusing here
;; would reject a valid component. The call side is
;; `test_calling_an_imported_component_func_needs_the_linker`.
;;
;; 21 × 2 = 42, so a lower that reached nothing returns 21 and one that dropped
;; its argument returns 0.

(component
  (core module $m
    (func (export "double") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "double" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))

  (import "host-op" (func $imported (type $ft)))
  (canon lift (core func $d) (func $lifted (type $ft)))

  (canon lower (func 1) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 21))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
