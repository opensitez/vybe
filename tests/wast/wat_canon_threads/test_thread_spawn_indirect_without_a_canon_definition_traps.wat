;; vybe-test: wast/wat_canon_threads/test_thread_spawn_indirect_without_a_canon_definition_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵② canon thread.spawn-indirect — "simply fuses the thread.new-indirect
;;   and thread.resume-later built-ins, allowing thread-creation to skip the
;;   intermediate suspended state transition".
;;
;; Because it IS `thread.new-indirect` plus `resume-later`, it inherits that
;; row's immediates: `$ftbl` is a canonical-definition immediate, not a stack
;; argument. Both rows share the same helper precisely so this cannot drift —
;; a spawn row that resolved its table differently from the new row would be
;; two implementations of one spec paragraph.

(module
  (import "canon" "thread.spawn-indirect" (func $spawn (param i32) (param i32) (result i32)))
  (table 1 funcref)
  (func (export "_start")
    i32.const 0
    i32.const 0
    call $spawn
    drop
  )
)
