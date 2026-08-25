;; vybe-test: wast/wat_component/test_instantiating_with_an_import_needs_the_linker
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md
;;   §Instance definitions — `instantiatearg ::= (with <name> <sortidx>)`
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   instance: `(with "f" (func 0))` supplies an IMPORT to a component
;;   instantiation, which needs the component linker (see cmplan.md §Deferred
;;   to export). A component with no imports instantiates without any `(with …)`
;;
;; ▶▶ INSTANTIATING WORKS; SUPPLYING AN IMPORT DOES NOT. The distinction is the
;; whole point of this file existing beside
;; `test_a_nested_component_runs_when_instantiated`, which instantiates the
;; same way and runs green. A `(with …)` clause binds one of the component's
;; IMPORTS, and an imported component function has no defining row anywhere in
;; the component — that is the single case that genuinely needs the linker.
;;
;; ⛔ ACCEPTING AND IGNORING THE CLAUSE IS THE FAILURE TO AVOID. The
;; instantiation would succeed with the import unsupplied, and the error would
;; surface later at the CALL, naming a missing callee rather than the clause
;; that was dropped. The refusal is at the clause because that is where the
;; information still exists.
;;
;; The message says what a caller can do instead — instantiate a component with
;; no imports — rather than only what it cannot.

(component
  (component $inner
    (type $ft (func (result u32)))
    (import "f" (func $needed (type $ft)))
  )
  (core module $m
    (func (export "zero") (result i32) (i32.const 0)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "zero" (core func $z))
  (type $ft2 (func (result u32)))
  (canon lift (core func $z) (func $mine (type $ft2)))

  (instance $i (instantiate $inner (with "f" (func $mine))))
)
