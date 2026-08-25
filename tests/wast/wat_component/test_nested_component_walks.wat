;; vybe-test: wast/wat_component/test_nested_component_walks
;; hand-written against proposals/component-model/design/mvp/Explainer.md §2
;;   `definition ::= … | <component> | …`
;;
;; A component is a definition of a component, so the grammar is recursive. The
;; discriminator matters: `componenttype ::= (component <componentdecl>*)` is
;; ALSO spelled `(component …)`, and it matches `(component)` with zero decls.
;; If the ordered choice resolved a nested definition into the TYPE position
;; instead, this would parse clean and walk nothing — `go` would not exist.
;;
;; ▶▶ THE INSTANTIATION IS NOW PART OF THE CLAIM, AND IT STRENGTHENS IT.
;; This file used to assert `go` existed with no `(instance …)` at all, because
;; a nested component was WALKED INLINE — its core modules ran where the
;; component was written. That was wrong: nothing inside a component should run
;; until something instantiates it, and a component instantiated twice would
;; still only have run once. Nested components are now DECLARED, like
;; `(core module …)` already was, and run at `(instance (instantiate …))`.
;;
;; So the grammar discriminator is sharper than before: a `componenttype` is
;; not instantiable, so if the ordered choice resolved `(component $c …)` into
;; the TYPE position, `(instantiate $c)` would refuse rather than quietly
;; producing nothing. The old shape could only detect the mis-parse by an
;; absence; this one detects it by a refusal.
;;
;; The nested component still has its OWN core module index space: `$m` here is
;; the inner one, and an outer `$m` would neither shadow it nor be visible to
;; it — pinned separately by
;; `test_nested_component_spaces_are_not_shared`.

(component
  (component $c
    (core module $m
      (func (export "go") (result i32)
        i32.const 7))
    (core instance (instantiate $m)))
  (instance (instantiate $c))
)
(assert_return (invoke "go") (i32.const 7))
