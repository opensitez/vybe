;; vybe-test: wast/wat_canon_threads/test_thread_suspend_outside_a_lifted_call_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.suspend:
;;       thread = current_thread()
;;       trap_if(not thread.task.inst.may_leave)
;;       cancelled = thread.suspend(cancellable)
;;
;; `current_thread()` is unconditional. A bare core module is not inside any
;; `canon lift`ed call, so there is no current thread to suspend — and answering
;; would mean suspending SOMEBODY, with index 0 naming another thread's slot.
;;
;; Blocking is the one thing separating this row from the four compound
;; handoffs: it passes no thread to `switch_to`.

(module
  (import "canon" "thread.suspend" (func $suspend (result i32)))
  (func (export "_start")
    call $suspend
    drop
  )
)
