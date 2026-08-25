;; vybe-test: wast/wat_call_tags/test_call_tag_type_must_be_supertype
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; "When defining a `func` of type `[ti*] -> [to*]`, one can optionally specify
;; `(call_tag $call_tag*)`, where each `$call_tag`'s type must be a supertype of
;; `[ti*] -> [to*]`."
;;
;; `$two_args` takes two i32 but claims a one-parameter tag, so the declaration
;; is invalid and the module must not run a call under it.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $one (param i32) (result i32))

  (func $two_args (param i32) (param i32) (result i32) (call_tag $one)
    local.get 0
  )

  (func (export "_start")
    i32.const 1
    ref.func $two_args
    call_with_tag $one
    call $log
  )
)
