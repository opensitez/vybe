;; vybe-test: wast/wat_call_tags/test_call_tag_unhandled_traps
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; "For canonical call tags, the answer is simply that the program traps."
;;
;; This is the property the whole proposal is for: a funcref called under a
;; convention it does not handle STOPS, rather than being called anyway under
;; the wrong shape. `$only` declares that it handles $t1 alone, so a call under
;; $t2 must not reach it.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (param i32) (result i32))
  (call_tag $t2 (param i32) (result i32))

  (func $only (param i32) (result i32) (call_tag $t1)
    local.get 0
  )

  (func (export "_start")
    i32.const 1
    ref.func $only
    call_with_tag $t2
    call $log
  )
)
