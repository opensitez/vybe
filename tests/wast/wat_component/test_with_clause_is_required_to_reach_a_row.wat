;; vybe-test: wast/wat_component/test_with_clause_is_required_to_reach_a_row
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §3
;;
;; The control for `test_with_clause_addresses_a_canon_row`. Byte-for-byte that
;; component with the `(with …)` clause DROPPED.
;;
;; Without it the import resolves from its own (module, name) pair, as any core
;; import does — `host:canon:spawn` — and "spawn" is not the name of a canonical
;; built-in. So the failure is about an unresolved callee, NOT the `shared`
;; refusal its twin reports.
;;
;; That difference is what makes the twin a proof: if a canon row could be
;; reached merely by declaring it in the same component, both files would trap
;; the same way and the `with` clause would be doing nothing.

(component
  (canon thread.spawn-ref shared (core type 0) (core func $sr))
  (core module $m
    (import "canon" "spawn"
      (func $spawn (param funcref) (param i32) (result i32)))
    (func (export "_start")
      ref.null func
      i32.const 0
      call $spawn
      drop))
  (core instance (instantiate $m))
)
