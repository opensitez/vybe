;; vybe-test: wast/wat_canon_threads/test_thread_suspend_then_resume_self_target_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.suspend-then-resume — `assert(self.running() and
;;   other.suspended())`.
;;
;; ⛔ POSITIVE THREAD PATH — unreachable from .wat until `canon lift` executed.
;; Every 🧵 row reaches `current_thread()` first, and a thread exists only
;; inside a lifted call, so this family could previously only be tested for the
;; trap it gives with NO thread at all.
;;
;; A thread hands off to ITSELF (`thread.index` supplies the target). The two
;; handoff families must answer this DIFFERENTLY, and both answers fall out of
;; the spec's own conditions rather than a special case:
;;
;;   resume  → self is running, therefore not suspended ⇒ TRAPS (here)
;;   promote → self is running, therefore not waiting, therefore not ready
;;             ⇒ falls back to plain suspend
;;             (test_thread_suspend_then_promote_self_target_blocks)
;;
;; An unconditional self-target trap — which is what this code used to have —
;; is wrong for the promote family, because it refuses a handoff the spec
;; DEFINES as a fallback. The pair of tests is what pins that the two families
;; disagree; either one alone would pass under the wrong implementation.
;;
;; `run-fail` is green on ANY failure, so the message MUST be read:
;;
;;   canon thread.suspend-then-resume: thread 0 is Running, must be Suspended (trap)

(component
  (core module $m
    (import "canon" "thread.suspend-then-resume" (func $h (param i32) (result i32)))
    (import "canon" "thread.index" (func $ti (result i32)))
    (func (export "run") (result i32)
      (call $h (call $ti))))
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
