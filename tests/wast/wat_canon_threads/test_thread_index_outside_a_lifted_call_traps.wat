;; vybe-test: wast/wat_canon_threads/test_thread_index_outside_a_lifted_call_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.index — `assert(thread.index is not None)`
;;
;; A CM thread exists only inside a `canon lift`ed call: the spec spawns the
;; IMPLICIT thread in `canon_lift` and there is no other way for one to come
;; into existence. A bare core module was never lifted, so `current_thread()`
;; has no answer.
;;
;; This must TRAP, not answer 0. Index 0 is a real slot in the instance's
;; `threads` table, so returning it would name another thread — and the caller
;; would have no way to tell a genuine thread 0 from "there is no thread".

(module
  (import "canon" "thread.index" (func $thread_index (result i32)))
  (func (export "_start")
    call $thread_index
    drop
  )
)
