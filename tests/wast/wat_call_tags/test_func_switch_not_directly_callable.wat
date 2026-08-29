;; vybe-test: wast/wat_call_tags/test_func_switch_not_directly_callable
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; Design §Functions: "`func_switch` is a new way of defining of functions …
;; The new function HAS NO TYPE AND CANNOT BE DIRECTLY CALLED, but we can get a
;; `funcref` for it by using `func.ref`, with the expectation that it later gets
;; called using `call_with_tag`."
;;
;; A `func_switch` has no type, so there is no signature for a direct `call` to
;; check against — which is exactly why the proposal forbids it rather than
;; leaving it to the type check. Reaching one through the front door selects no
;; arm and answers whatever the first arm happens to be.
;;
;; The CONTROL for this test is `test_func_switch_dispatch`: the same switch,
;; reached the sanctioned way through `call_with_tag`, must keep working. A
;; rejection that also broke that would be a worse bug than the one being fixed.
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t (canon) (param i32) (result i32))

  (func $arm (param i32) (result i32)
    local.get 0
  )

  (func_switch $fs (case $t $arm))

  (func (export "_start")
    i32.const 5
    call $fs
    call $log
  )
)
