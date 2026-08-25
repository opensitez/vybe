;; vybe-test: wast/wat_component/test_inline_lift_type_refuses
;; vybe-test-mode: compile-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §9
;;   `canon ::= (canon lift core-prefix(<core:funcsortidx>) <canonopt>*
;;               bind-id(<externtype>))`
;;
;; The lifted type may be written inline. The canon SECTION, though, records an
;; `ft:<typeidx>` — an index — and an inline type has no index to record.
;;
;; So this refuses, naming the fix. The alternative was to record `None` and
;; let a downstream `require_type` report a missing `$ft` immediate — which
;; would be a lie: the source supplied the type, the front end dropped it.
;; An absent immediate and a discarded one must not produce the same message.

(component
  (core module (func (export "go") (result i32) i32.const 1))
  (canon lift (core func 0) (func (result u32)))
)
