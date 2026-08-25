;; vybe-test: wast/wat_component/test_unbound_canon_type_id_refuses
;; vybe-test-mode: compile-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §1
;;   `idx ::= <u32> | <core:id>`
;;
;; `$nope` names nothing in the component's type space. The refusal must name
;; the binding and the SPACE:
;;
;;   canon: `$nope` is not bound in the component type index space
;;
;; Resolving an unbound `$id` to 0 is the `GLOBAL_GET` defect with a new pair
;; of tables — it links, it runs, and it addresses the wrong definition with
;; nothing downstream able to detect it. Each index space carries its own name
;; map for the same reason: `(core type $t)` and `(core func $t)` may both be
;; bound, to different entities, and one shared map would silently answer one
;; for the other.

(component
  (canon stream.read $nope (core func $sr))
)
