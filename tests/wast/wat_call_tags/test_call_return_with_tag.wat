;; vybe-test: wast/wat_call_tags/test_call_return_with_tag
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "`call_return_with_tag $call_tag : [ti* funcref] -> [to*]` […] tail calls the
;; given `funcref` with the specified call tag."
;;
;; Same tag semantics as `call_with_tag`; it differs only in not growing the
;; frame where the engine can avoid it, so the observable result is identical.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t (canon) (param i32) (result i32))

  (func $negate (param i32) (result i32)
    i32.const 0
    local.get 0
    i32.sub
  )

  (func $tail (param i32) (result i32)
    local.get 0
    ref.func $negate
    call_return_with_tag $t
  )

  (func (export "_start")
    i32.const 42
    call $tail
    call $log
  )
)
