;; vybe-test: wast/wat_call_tags/test_func_switch_no_forward_uses_tag_fallback
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "If there is no corresponding call tag, then if `$func_switch` is specified
;;  the call tag and arguments are forwarded to it, OTHERWISE THE FALL-BACK
;;  HANDLER OF THE CALL TAG IS (TAIL) CALLED with the arguments."
;;
;; The func_switch grammar has THREE outcomes and only two were covered:
;;   match                      -> test_func_switch_dispatch
;;   no match + (forward $fs)   -> test_func_switch_forward
;;   no match + NO forward      -> THIS TEST — the CALL TAG's fallback runs
;;
;; $sw handles $t1 only and specifies no forward. Called under $t2, there is no
;; case and nothing to forward to, so $t2's own fall-back handler answers — and
;; it receives the refusing funcref as its trailing parameter, per
;; `[ti* funcref] -> [to*]`.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (param i32) (result i32))
  (call_tag $t2 (param i32) (result i32) (fallback $on_miss))

  (func $on_miss (param i32) (param externref) (result i32)
    i32.const 55
  )

  (func $t1_impl (param i32) (result i32)
    local.get 0
  )

  (func_switch $sw
    (case $t1 $t1_impl)
  )

  (func (export "_start")
    i32.const 7
    ref.func $sw
    call_with_tag $t2
    call $log
  )
)
