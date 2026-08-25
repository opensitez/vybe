;; vybe-test: wast/wat_call_tags/test_call_tag_multiple_tags_one_func
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "These can be canonical call tags for *multiple* type signatures […] Notice
;; that a `funcref` can handle multiple call tags."
;;
;; The signature-refinement case: `$impl` declares BOTH tags, so a caller naming
;; either reaches it — which is how the proposal avoids needing structural
;; subtyping of typed function references.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $super (param i32) (result i32))
  (call_tag $sub (param i32) (result i32))

  (func $impl (param i32) (result i32) (call_tag $super $sub)
    local.get 0
    i32.const 1
    i32.add
  )

  (func (export "_start")
    i32.const 1
    ref.func $impl
    call_with_tag $super
    call $log

    i32.const 2
    ref.func $impl
    call_with_tag $sub
    call $log
  )
)
