;; vybe-test: wast/wat_call_tags/test_call_tag_fallback_arity
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; "This `$func` must have the same signature as `$functype` *except* also
;; accepting an additional `funcref` so that we can pass the fall-back handler
;; the specific `funcref` that did *not* recognize the call tag."
;;
;; `$bad_handler` omits the trailing funcref, so it cannot be a handler for a
;; one-parameter tag.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t (param i32) (result i32) (fallback $bad_handler))

  (func $bad_handler (param i32) (result i32)
    i32.const 0
  )

  (func $other (param i32) (result i32)
    local.get 0
  )

  (func (export "_start")
    i32.const 1
    ref.func $other
    call_with_tag $t
    call $log
  )
)
