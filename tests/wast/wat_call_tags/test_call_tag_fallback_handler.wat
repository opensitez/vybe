;; vybe-test: wast/wat_call_tags/test_call_tag_fallback_handler
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "when creating call tags with `call_tag.new $functype`, one can also specify
;; a `$func` to use as its fall-back handler. This `$func` must have the same
;; signature as `$functype` *except* also accepting an additional `funcref` so
;; that we can pass the fall-back handler the specific `funcref` that did *not*
;; recognize the call tag."
;;
;; `$only` handles $t1, so calling it under $t2 misses — and $t2 has a handler,
;; so the handler runs instead of trapping. It answers 99 to prove the fall-back
;; result is what reaches the call site.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (param i32) (result i32))
  (call_tag $t2 (param i32) (result i32) (fallback $on_miss))

  ;; `[ti* funcref] -> [to*]` — the trailing funcref is the one that refused.
  (func $on_miss (param i32) (param externref) (result i32)
    i32.const 99
  )

  (func $only (param i32) (result i32) (call_tag $t1)
    local.get 0
  )

  (func (export "_start")
    i32.const 7
    ref.func $only
    call_with_tag $t2
    call $log
  )
)
