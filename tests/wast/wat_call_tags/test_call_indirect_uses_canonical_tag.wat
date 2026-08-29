;; vybe-test: wast/wat_call_tags/test_call_indirect_uses_canonical_tag
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; "(And `call_indirect $table $functype` is now shorthand for
;;  `(call_with_tag (call_tag.canon $functype) (table.get $table))`.)"
;;
;; combined with:
;;
;; "the `funcref` returned by `func.ref` for this `func` handles exactly the
;;  call tags in `$call_tag*`."
;;
;; $only declares $t1 and therefore handles EXACTLY $t1 — NOT the canonical tag
;; of its own signature. So the plain `call_indirect` below is a call under a
;; convention it does not handle, and must trap.
;;
;; This is the Security property stated in §Applications: "a funcref called
;; under a convention it does not handle STOPS, rather than being called anyway
;; under the wrong shape." Checking it only in `call_with_tag` leaves the front
;; door open — `call_indirect` IS a call_with_tag by definition.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (param i32) (result i32))
  (type $ft (func (param i32) (result i32)))

  (func $only (param i32) (result i32) (call_tag $t1)
    local.get 0
  )

  (table 1 funcref)
  (elem (i32.const 0) $only)

  (func (export "_start")
    i32.const 7
    i32.const 0
    call_indirect (type $ft)
    call $log
  )
)
