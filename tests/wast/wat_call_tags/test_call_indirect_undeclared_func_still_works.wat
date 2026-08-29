;; vybe-test: wast/wat_call_tags/test_call_indirect_undeclared_func_still_works
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "By default, this `funcref` handles the canonical call tag for the function's
;;  signature."
;;
;; The CONTROL for `test_call_indirect_uses_canonical_tag`. A func that declares
;; no tags keeps the canonical tag of its own signature, so plain
;; `call_indirect` — defined as `call_with_tag (call_tag.canon $functype)` —
;; must still reach it unchanged.
;;
;; Without this arm the negative test above proves nothing: a canonical-tag
;; check that rejected EVERY indirect call would pass it just as well.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (type $ft (func (param i32) (result i32)))

  (func $double (param i32) (result i32)
    local.get 0
    i32.const 2
    i32.mul
  )

  (table 1 funcref)
  (elem (i32.const 0) $double)

  (func (export "_start")
    i32.const 21
    i32.const 0
    call_indirect (type $ft)
    call $log
  )
)
