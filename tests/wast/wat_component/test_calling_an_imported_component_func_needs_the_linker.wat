;; vybe-test: wast/wat_component/test_calling_an_imported_component_func_needs_the_linker
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2601
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   canon lower: component func 0 is IMPORTED — it has no defining row in this
;;   component, so calling it needs the component linker (see cmplan.md
;;   §Deferred to export)
;;
;; ▶▶ THE ONE PLACE THE COMPONENT LINKER IS GENUINELY REQUIRED. `canon lower`
;; was said for months to need the linker; it does not — see
;; `test_canon_lower_of_a_lifted_function`, where `canon_lower(callee, ft, …)`
;; takes `ft` from the callee and the row carries no `ft` immediate at all.
;; What lower needed was the component FUNCTION INDEX SPACE. This file is the
;; remainder: an IMPORTED component function has no defining row anywhere in
;; this component, so nothing local can supply the callee.
;;
;; ⛔ IT REFUSES AT THE CALL, NOT AT THE DECLARATION. A component may
;; legitimately import something it never uses;
;; `test_an_import_occupies_a_component_func_index` declares this same import
;; and runs green. Moving the refusal earlier would reject valid components,
;; and moving it later — to a default callee — would call something arbitrary.
;;
;; The message must stay DISTINCT from the out-of-range one. "Slot empty" and
;; "no such slot" are different mistakes: a callee nobody supplied versus a
;; stale index. Collapsing them is how a linker gap starts reading as a typo.

(component
  (type $ft (func (param "a" u32) (result u32)))
  (import "host-op" (func $imported (type $ft)))

  (canon lower (func $imported) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "_start")
      (drop (call $l (i32.const 21)))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)
