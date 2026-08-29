;; vybe-test: wast/wat_call_tags/test_call_tag_fallback_funcref_is_the_refuser
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; Design §Implementation: "If the call tag is not recognized, then the code
;; jumps-to/tail-calls the fall-back handler pointed to by the call tag, leaving
;; all the arguments in their place but replacing the call-tag value with the
;; value of the current `funcref`."
;;
;; ⛔ THE EXISTING FALL-BACK TEST CANNOT SEE THIS. `test_call_tag_fallback_handler`
;; declares `$on_miss (param i32) (param externref)` and then IGNORES BOTH,
;; returning a constant — so it proves the handler is reached with the right
;; ARITY and nothing about WHICH funcref arrives. A null, a wrong funcref or
;; garbage in that slot all pass it.
;;
;; This one READS the parameter: the handler calls the funcref it was handed,
;; under a tag that funcref DOES declare. Only the genuine `$only` answers 107,
;; so the assertion below can only hold if the refusing funcref itself arrived.
;; The check is a trap rather than a logged value, so it holds whatever the
;; harness compares.
;;
;; This clause is load-bearing beyond wast: `primitives/dynamic_invoke.rs` rests
;; its correctness argument on it — "case (2) carrying the funcref is what makes
;; stamping an OPTIMISATION rather than a correctness requirement: a method that
;; exists but was not stamped for this arity arrives at the handler as a live
;; funcref and is simply called."
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (param i32) (result i32))
  (call_tag $t2 (param i32) (result i32) (fallback $on_miss))

  ;; Handles $t1 only, so a $t2 call misses it.
  (func $only (param i32) (result i32) (call_tag $t1)
    local.get 0
    i32.const 100
    i32.add
  )

  ;; `[ti* funcref] -> [to*]` — the trailing funcref is the one that refused.
  (func $on_miss (param i32) (param externref) (result i32)
    local.get 0
    local.get 1
    call_with_tag $t1
  )

  (func (export "_start")
    i32.const 7
    ref.func $only
    call_with_tag $t2
    i32.const 107
    i32.ne
    if
      unreachable
    end
    i32.const 1
    call $log
  )
)
