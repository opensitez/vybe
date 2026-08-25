;; vybe-test: wast/wat_canon_threads/test_thread_yield_outside_a_lifted_call_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🔀 canon thread.yield:
;;       def canon_thread_yield(cancellable):
;;         thread = current_thread()
;;         trap_if(not thread.task.inst.may_leave)
;;         cancelled = thread.yield_(cancellable)
;;         return [cancelled]
;;
;; ⛔ This test USED to assert the opposite — that `thread.yield` succeeds in a
;; bare core module — on the claim that it is the one 🧵-adjacent row needing no
;; current thread. That claim was wrong. `current_thread()` is unconditional
;; here, exactly as in `thread.index`, and the old test was pinning a deviation
;; rather than the spec.
;;
;; Core wasm inside a real component is always inside a lifted call, so this
;; situation cannot arise in a valid component at all — a bare core module run
;; standalone is already outside the component model. Trapping says that; the
;; previous silent success said the guest had a thread when it did not.

(module
  (import "canon" "thread.yield" (func $thread_yield (result i32)))
  (func (export "_start")
    call $thread_yield
    drop
  )
)
