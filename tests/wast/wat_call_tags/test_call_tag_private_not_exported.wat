;; vybe-test: wast/wat_call_tags/test_call_tag_private_not_exported
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "rather than using canonical call tags, the application can use solely custom
;; non-exported call tags. This guarantees that the only calls to a function
;; that can be made through function references comes from `call_with_tag`
;; instructions within the module."
;;
;; `$private` is never exported, and `$guarded` handles only it — so the func is
;; reachable from inside the module and by nothing outside. The call below is
;; that inside call, and must succeed.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $private (param i32) (result i32))

  (func $guarded (param i32) (result i32) (call_tag $private)
    local.get 0
    i32.const 7
    i32.mul
  )

  (func (export "_start")
    i32.const 6
    ref.func $guarded
    call_with_tag $private
    call $log
  )
)
