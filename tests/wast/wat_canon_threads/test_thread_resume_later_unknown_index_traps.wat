;; vybe-test: wast/wat_canon_threads/test_thread_resume_later_unknown_index_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.resume-later:
;;       other_thread = inst.threads.get(i)
;;       trap_if(not other_thread.suspended())
;;
;; `inst.threads.get(i)` is a table lookup that must FIND something. A module
;; that never created a thread has an empty table, so every index is absent.
;;
;; Answering silently would be worse than useless: `resume-later` marks a thread
;; ready to run at some later point chosen by the embedder, so a no-op looks
;; exactly like a thread that simply has not been scheduled yet, and the guest
;; would wait forever for something it never actually queued.

(module
  (import "canon" "thread.resume-later" (func $resume_later (param i32)))
  (func (export "_start")
    i32.const 0
    call $resume_later
  )
)
