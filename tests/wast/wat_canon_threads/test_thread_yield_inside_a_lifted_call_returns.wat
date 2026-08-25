;; vybe-test: wast/wat_canon_threads/test_thread_yield_inside_a_lifted_call_returns
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🔀 canon thread.yield:
;;
;;     def canon_thread_yield(cancellable):
;;       thread = current_thread()
;;       trap_if(not thread.task.inst.may_leave)
;;       cancelled = thread.yield_(cancellable)
;;       return [cancelled]
;;
;; ⛔ POSITIVE THREAD PATH — unreachable from .wat until `canon lift` executed.
;;
;; `current_thread()` is UNCONDITIONAL. An earlier comment in the runtime
;; claimed yield needs no current thread, and its test ASSERTED that as
;; correct; the test has since changed sign. Its negative half is
;; test_thread_yield_outside_a_lifted_call_traps.
;;
;; This is the half that could never run. `yield_` is `wait_until(lambda: True)`
;; — the condition already holds, so it returns rather than switching, and
;; `cancelled` is 0 because no cancellation was requested. That 0 is the
;; assert, and it bites: expecting 1 fails.

(component
  (core module $m
    (import "canon" "thread.yield" (func $y (result i32)))
    (func (export "run") (result i32)
      (call $y)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "run" (core func $r))

  (type $ft (func (result u32)))
  (canon lift (core func $r) (func $lifted (type $ft)))

  (core module $caller
    (import "canon" "lift@0" (func $l (result i32)))
    (func (export "get") (result i32)
      (call $l)))
  (core instance (instantiate $caller))
)

(assert_return (invoke "get") (i32.const 0))
