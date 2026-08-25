;; vybe-test: wast/wat_component/test_canon_immediate_survives_the_carrier
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Binary.md §296
;;   `(canon thread.spawn-ref shared? <core:typeidx> (core func <id>?))` 🧵②
;;
;; ⛔ The companion to `test_canon_section_reaches_the_vm`, and the one that
;; matters more.
;;
;; That test proves the section arrives with the right LENGTH (`have 2`). It
;; does NOT prove any FIELD survived walker → `Module::canon` →
;; `compiler::canon::lower_section` → `Chunk::canon_section` →
;; `VM::merge_canon_section`. `lower_section` could drop `cancellable`,
;; mis-map `string-encoding`, or swap `ty`/`table` on the spawn rows, and every
;; other test here would still be green.
;;
;; `shared?` is the field to prove it with, because `refuse_shared_threads`
;; reads `canon_defs[canonidx].shared` and traps on it. The trap message is the
;; discriminator and MUST be read — this is a `run-fail` test, so it is green on
;; ANY failure, and the null funcref below fails on its own:
;;
;;   canon thread.spawn-ref: `shared` requests a PREEMPTIVE thread ... (trap)
;;
;; NOT `canon thread.spawn-ref: null funcref (trap)`, which is what BOTH of the
;; controls report:
;;
;;   * the same core module with no component around it — `canon_defs` is empty,
;;     so `shared` cannot exist (`wat_canon_threads`'
;;     `test_thread_spawn_ref_null_funcref_traps` is that module);
;;   * the same component with `shared` DROPPED from the canon row — which is
;;     what makes this a proof and not a coincidence. If merely having a canon
;;     section produced the trap, that control would trap too. It does not.
;;
;; The `(core instance …)` is what runs the module — a component's core
;; module is DECLARED, not instantiated implicitly.
;;
;; `available-parallelism` cannot be used for this: it spawns nothing and
;; deliberately does not refuse `shared`.

(component
  (canon thread.spawn-ref shared (core type 0) (core func $sr))
  (core module $m
    (import "canon" "thread.spawn-ref@0"
      (func $spawn (param funcref) (param i32) (result i32)))
    (func (export "_start")
      ref.null func
      i32.const 0
      call $spawn
      drop))
  (core instance (instantiate $m))
)
