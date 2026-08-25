;; vybe-test: wast/wat_component/test_canon_section_reaches_the_vm
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Binary.md §296
;;   "Canonical Definitions" — a section of typed definition RECORDS.
;;
;; ⛔ This is the test that `VM::canon_defs` has a PRODUCER.
;;
;; The field was declared, read in nine places in `dispatch.rs`, and written by
;; NOTHING. Every canon row in the tree therefore fell through to the identity
;; section, where canonidx is read as a typeidx — which meant `cancellable?`
;; was false on every row and no `$t` or `opts` immediate ever arrived.
;;
;; The trap message is the discriminator, and it must be READ, not merely
;; observed to be a failure. It names the section's LENGTH:
;;
;;   canon thread.new-indirect: canonidx 3 is not a row of `VM::canon_defs`
;;   (have 2)
;;
;; `have 2` is the whole point. The companion test in `wat_canon_threads`
;; (`test_thread_new_indirect_unknown_canonidx_traps`) is byte-for-byte this
;; core module with NO component around it, and it reports `have 0`. Two canon
;; rows are declared below; if the section did not reach the VM — walker → AST
;; `Module::canon` → `compiler::canon::lower_section` → `Chunk::canon_section`
;; → `VM::merge_canon_section` — this would say `have 0` and still be green,
;; because a `run-fail` test passes on ANY failure.
;;
;; The `(core instance …)` is what runs the module — a component's core module
;; is DECLARED, not instantiated implicitly.

(component
  (canon backpressure.inc (core func $a))
  (canon backpressure.dec (core func $b))
  (core module $m
    (import "canon" "thread.new-indirect@3" (func $tn (param i32 i32) (result i32)))
    (table 1 funcref)
    (func (export "_start")
      i32.const 0
      i32.const 0
      call $tn
      drop))
  (core instance (instantiate $m))
)
