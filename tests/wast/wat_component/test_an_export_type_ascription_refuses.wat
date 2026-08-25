;; vybe-test: wast/wat_component/test_an_export_type_ascription_refuses
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2611
;;   export ::= (export <id>? <externnamelit> <attribute>* <externidx> <externtype>?)
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   export: `(export "run" (func $lifted) (func (result u32)))` — the trailing
;;   `<externtype>` ascribes a type to the export, and nothing checks it
;;   against the item being exported
;;
;; ▶▶ THE ASCRIPTION IS A CLAIM, AND AN UNCHECKED CLAIM IS WORSE THAN A
;; REFUSAL. `(export "run" (func $f) <externtype>)` narrows what the component
;; publishes for `$f` — the spec uses it to hide a resource type or to export a
;; subtype. Accepting it and ignoring it would report a type the export was
;; never proven to have, and the error would surface in the CONSUMER, which by
;; then has no way to know the ascription was never applied.
;;
;; The ascription here is deliberately WRONG — `$lifted` takes a parameter and
;; the ascribed type takes none — so if a later change starts honouring
;; ascriptions, this file must go on failing, with a mismatch message instead
;; of a "nothing checks it" one. Rewrite the header then rather than deleting
;; the file: the two refusals are different claims.

(component
  (core module $m
    (func (export "double") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "double" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))

  (export "run" (func $lifted) (func (result u32)))
)
