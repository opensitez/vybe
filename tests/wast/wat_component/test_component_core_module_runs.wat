;; vybe-test: wast/wat_component/test_component_core_module_runs
;; hand-written against proposals/component-model/design/mvp/Explainer.md §2, §3
;;   `definition        ::= core-prefix(<core:module>) | core-prefix(<core:instance>) | …`
;;   `core:instanceexpr ::= (instantiate <core:moduleidx> <core:instantiatearg>*)`
;;
;; `languages/wast` parsed MODULE SYNTAX ONLY until this landed
;; (`grammar.pest:4`), so a `(component …)` was a parse error.
;;
;; The core module is DECLARED by `(core module …)` and does not run; the
;; `(core instance (instantiate $m))` is what instantiates it. That ordering is
;; not a detail — a module's imports cannot be resolved until the instantiation
;; that supplies them has been read, and the instantiation comes AFTER the
;; module. Instantiating eagerly at the declaration would make the `with` clause
;; unimplementable, which is exactly the mistake this shape avoids.

(component
  (core module $m
    (func (export "go") (result i32)
      i32.const 42))
  (core instance $i (instantiate $m))
)

(assert_return (invoke "go") (i32.const 42))
