;; vybe-test: wast/wat_canon_threads/test_thread_spawn_ref_null_funcref_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵② canon thread.spawn-ref — validation gives $spawn_ref the type
;;   `(func (param $f (ref null $ft)) (param $c T) (result $e i32))`.
;;
;; `ref null $ft` admits a NULL reference at the type level, so the null has to
;; be rejected where the thread is created rather than by validation. There is
;; no body to run and no honest thread index to return for one.

(module
  (import "canon" "thread.spawn-ref" (func $spawn (param funcref) (param i32) (result i32)))
  (func (export "_start")
    ref.null func
    i32.const 0
    call $spawn
    drop
  )
)
