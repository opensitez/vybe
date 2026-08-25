;; vybe-test: wast/wat_canon_threads/test_thread_new_indirect_without_a_canon_definition_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §🧵 canon thread.new-indirect:
;;       f = ftbl.get(fi)
;;       trap_if(f.t != ft)
;;
;; `$ft` and `$ftbl` are IMMEDIATES of the canonical definition
;; `(canon thread.new-indirect $ft $ftbl)` — not runtime arguments. Only `fi`
;; and the closure `c` come off the stack.
;;
;; So a call site that carries no canonical definition is not an under-specified
;; call, it is an unrepresentable one: there is no table to look `fi` up in and
;; no type to check the funcref against. Trapping is the only honest answer;
;; picking table 0 by default would silently thread through whichever table the
;; module happened to define and check nothing.

(module
  (import "canon" "thread.new-indirect" (func $thread_new (param i32 i32) (result i32)))
  (table 1 funcref)
  (func (export "_start")
    i32.const 0   ;; fi
    i32.const 0   ;; c
    call $thread_new
    drop
  )
)
