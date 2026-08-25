;; vybe-test: wast/wat_call_tags/test_call_tag_canonical_default
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "By default, this `funcref` handles the canonical call tag for the function's
;; signature." A func that declares no tags stays reachable through the
;; canonical tag of its own type — which is what keeps `call_indirect`
;; (defined as `call_with_tag (call_tag.canon $functype)`) working unchanged.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (canon) (param i32) (result i32))

  (func $double (param i32) (result i32)
    local.get 0
    i32.const 2
    i32.mul
  )

  (func (export "_start")
    i32.const 21
    ref.func $double
    call_with_tag $t1
    call $log
  )
)
