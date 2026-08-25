;; vybe-test: wast/wat_canon_threads/test_thread_available_parallelism_is_one
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵② canon thread.available-parallelism — "returns the number of threads
;;   the underlying hardware can be expected to execute in parallel ... not
;;   allowed to change over the lifetime of a component instance".
;;
;; This runtime schedules COOPERATIVE fibers: exactly one thread runs at a time,
;; so the honest count is 1. That is not the deterministic profile's `return [1]`
;; standing in for a real number — it IS the real number for this scheduler.
;;
;; Called twice because the spec's "not allowed to change" is a property of the
;; SEQUENCE, not of any single call, and a row that read a mutable counter would
;; satisfy a one-call test.

(module
  (import "canon" "thread.available-parallelism" (func $par (result i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
    call $par
    i32.const 1 call $vybe_check_i32

    call $par
    i32.const 1 call $vybe_check_i32
  )
)
