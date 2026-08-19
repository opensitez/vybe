;; vybe-test: wast/wat_canon_stream/test_resource_new_rep_drop
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon resource.new  — (func (param $rt.rep) (result i32))
;;   §canon resource.rep  — (func (param i32) (result $rt.rep))
;;   §canon resource.drop — (func (param i32))
;;
;; A resource handle is an INDIRECTION: the component keeps a private
;; representation and hands out an opaque index. `resource.rep` is the only way
;; back, and it is valid only for a handle of the same resource type — a peer
;; holding your handle can never read another type's representation through it.
;;
;; `@7` is the `$rt` immediate: these are all the same resource type.

(module
  (import "canon" "resource.new@7" (func $res_new (param i32) (result i32)))
  (import "canon" "resource.rep@7" (func $res_rep (param i32) (result i32)))
  (import "canon" "resource.drop@7" (func $res_drop (param i32)))
  (memory 1)
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start") (local $h i32) (local $h2 i32)
    ;; wrap a representation
    i32.const 1234
    call $res_new
    local.set $h

    ;; the handle is NOT the representation — that is the whole point
    local.get $h
    i32.const 1234
    i32.eq
    if
      unreachable
    end

    ;; and it maps back
    local.get $h
    call $res_rep
    i32.const 1234 call $vybe_check_i32

    ;; a second resource gets a distinct handle
    i32.const 5678
    call $res_new
    local.set $h2
    local.get $h
    local.get $h2
    i32.eq
    if
      unreachable
    end
    local.get $h2
    call $res_rep
    i32.const 5678 call $vybe_check_i32

    ;; dropping one leaves the other intact
    local.get $h
    call $res_drop
    local.get $h2
    call $res_rep
    i32.const 5678 call $vybe_check_i32
  )
)
