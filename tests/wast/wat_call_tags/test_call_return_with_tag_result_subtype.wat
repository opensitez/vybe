;; vybe-test: wast/wat_call_tags/test_call_return_with_tag_result_subtype
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; Design §Instructions: "`call_return_with_tag $call_tag : [ti* funcref] -> [to*]`
;; … tail calls the given `funcref` … WHERE `[to*]` IS A SUBTYPE OF THE RESULT
;; TYPE OF THE FUNCTION CONTAINING THIS INSTRUCTION."
;;
;; A tail call returns the callee's results directly to THIS function's caller,
;; so the callee's results have to be something this function was allowed to
;; return. Here `$t` yields `[i32]` while `$bad` declares no results at all —
;; `[i32]` is not a subtype of `[]` — so the module must be rejected.
;;
;; The CONTROL is `test_call_return_with_tag`, where the tag's results and the
;; containing function's results agree and the call must still work.
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t (canon) (param i32) (result i32))

  (func $target (param i32) (result i32)
    local.get 0
  )

  ;; Declares NO results, but tail-calls a tag whose results are [i32].
  (func $bad (param i32)
    local.get 0
    ref.func $target
    call_return_with_tag $t
  )

  (func (export "_start")
    i32.const 5
    call $bad
    i32.const 0
    call $log
  )
)
