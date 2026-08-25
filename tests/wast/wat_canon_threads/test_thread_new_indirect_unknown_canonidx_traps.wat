;; vybe-test: wast/wat_canon_threads/test_thread_new_indirect_unknown_canonidx_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Binary.md §296
;;   `canon ::= ... thread.new-indirect <typeidx> <tableidx>`
;;
;; The companion to the bare-name case: `@3` DOES name a canonical definition,
;; it just names one this module never declared. A core module carries no canon
;; section at all, so every canonidx is absent and index 3 resolves to nothing.
;;
;; This is the failure the `@N` convention exists to make loud. An out-of-range
;; canonidx that silently fell back to row 0 would run the thread against some
;; other definition's table and type — the same one-integer-two-index-spaces
;; defect that made `global.get` read a constant instead of a slot.

(module
  (import "canon" "thread.new-indirect@3" (func $thread_new (param i32 i32) (result i32)))
  (table 1 funcref)
  (func (export "_start")
    i32.const 0
    i32.const 0
    call $thread_new
    drop
  )
)
