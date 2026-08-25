;; vybe-test: wast/wat_component/test_with_clause_wires_one_instance_to_another
;; hand-written against proposals/component-model/design/mvp/Explainer.md §3
;;   `core:instantiatearg ::= (with <core:name> (instance <core:instanceidx>))`
;;
;; Whole-instance wiring: every export of `$li` fills the "lib" slot under its
;; own name. This is how a core module inside a component imports from another
;; core module, and it needs THREE things that did not exist:
;;
;;   * a core INSTANCE index space, so `$li` names something;
;;   * that instance's export table, so "answer" resolves;
;;   * the instantiation's wiring beating the `(register "m")` table, so the
;;     link is the component's and not the script's.
;;
;; Note there is no `(register …)` anywhere — that is the point. The script
;; level never learns "lib" exists; the component wires it.
;;
;; The assert is the discriminator: 42 can only arrive by actually calling
;; through the link. An unresolved import would trap instead.

(component
  (core module $lib
    (func (export "answer") (result i32)
      (i32.const 42)))
  (core instance $li (instantiate $lib))

  (core module $main
    (import "lib" "answer" (func $answer (result i32)))
    (func (export "get") (result i32)
      (call $answer)))
  (core instance (instantiate $main
    (with "lib" (instance $li))))
)

(assert_return (invoke "get") (i32.const 42))
