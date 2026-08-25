;; vybe-test: wast/wat_canon_threads/test_thread_suspend_then_promote_self_target_blocks
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.suspend-then-promote — `assert(self.running())` then
;;   `if other.ready(): ... else: <plain suspend>`.
;;
;; ⛔ POSITIVE THREAD PATH — unreachable from .wat until `canon lift` executed.
;;
;; ▶▶ THIS PINS **TWO** DEFECTS AT ONCE, both previously corrected-but-unverified:
;;
;;  1. `suspend-then-promote` was a NO-OP whenever its target was not ready —
;;     the fallback returned 0. The spec's fallback is plain `suspend`, which
;;     BLOCKS, and unlike `yield_` it has no early return. If the old behaviour
;;     came back, `get` would return 0 and the assert below would pass
;;     SILENTLY; the trap is the only thing that distinguishes them.
;;  2. The self-target trap was wrong for the PROMOTE family. Self is running,
;;     therefore not waiting, therefore not ready ⇒ the else branch. Its
;;     control is test_thread_suspend_then_resume_self_target_traps, which is
;;     the same source with `resume` and DOES trap on the self-target.
;;
;; So the message here must be the DEADLOCK, not a self-target refusal and not
;; a "must be Suspended". `run-fail` is green on ANY failure — read it:
;;
;;   canon thread.suspend-then-promote: no thread is ready and no host work is
;;   pending — this thread blocks with nothing left that could ever wake it (trap)
;;
;; That message is the correct answer here: the fallback genuinely blocked, and
;; this program has no other thread and no pending host work, which is one of
;; the three distinct outcomes `thread.suspend` must tell apart.

(component
  (core module $m
    (import "canon" "thread.suspend-then-promote" (func $h (param i32) (result i32)))
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
