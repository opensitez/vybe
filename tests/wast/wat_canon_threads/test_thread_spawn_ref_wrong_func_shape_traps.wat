;; vybe-test: wast/wat_canon_threads/test_thread_spawn_ref_wrong_func_shape_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵② canon thread.spawn-ref — "$ft must refer to the type
;;   `(shared? (func (param $c T)))` where T is i32".
;;
;; The thread function takes the closure and returns NOTHING. `$nullary` below
;; takes no parameter, so `[c]` would have nowhere to go — the closure would be
;; silently discarded and the thread would run with no argument at all.
;;
;; Checked at the same fidelity `call_indirect` uses for its `(type $sig)`: the
;; param and result COUNTS on the callee. The VM is untyped, so this is not a
;; structural comparison and does not claim to be — but `(func (param i32))` is
;; fully determined by its two counts, so nothing is lost here.

(module
  (import "canon" "thread.spawn-ref" (func $spawn (param funcref) (param i32) (result i32)))
  (elem declare func $nullary)
  (func $nullary)
  (func (export "_start")
    ref.func $nullary
    i32.const 0
    call $spawn
    drop
  )
)
