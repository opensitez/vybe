;; vybe-test: wast/wat_canon_threads/test_thread_handoff_outside_a_lifted_call_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.suspend-then-resume:
;;       thread = current_thread()
;;       other_thread = thread.task.inst.threads.get(i)
;;       trap_if(not other_thread.suspended())
;;
;; The compound handoffs need BOTH ends: a current thread to park, and a target
;; to enter. Outside a lifted call there is no current thread, so there is
;; nothing to hand off FROM — the built-in cannot fall back to "just run the
;; target", because the whole operation is defined as a transfer.

(module
  (import "canon" "thread.suspend-then-resume" (func $handoff (param i32) (result i32)))
  (func (export "_start")
    i32.const 0
    call $handoff
    drop
  )
)
