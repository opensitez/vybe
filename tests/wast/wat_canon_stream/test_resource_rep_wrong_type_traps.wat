;; vybe-test: wast/wat_canon_stream/test_resource_rep_wrong_type_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon resource.rep — `trap_if(h.rt is not rt)`
;;
;; A handle carries its resource TYPE, and reading the representation out of a
;; handle belonging to a different type is a trap. Without that check, holding
;; any handle would let a component read the private representation of every
;; other resource type it could name — the indirection would protect nothing.

(module
  (import "canon" "resource.new@7" (func $res_new_7 (param i32) (result i32)))
  (import "canon" "resource.rep@9" (func $res_rep_9 (param i32) (result i32)))
  (memory 1)
  (func (export "_start") (local $h i32)
    i32.const 1234
    call $res_new_7
    local.set $h

    ;; made as type 7, read as type 9
    local.get $h
    call $res_rep_9
    drop
  )
)
