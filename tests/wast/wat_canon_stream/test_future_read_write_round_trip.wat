;; vybe-test: wast/wat_canon_stream/test_future_read_write_round_trip
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon {stream,future}.new  — (func (result i64)), [ ri | (wi << 32) ]
;;   §canon future.{read,write}  — (func (param i32 T) (result i32))
;;
;; A future carries EXACTLY ONE element, which is why `future.read` takes a
;; pointer and NO count: the spec fixes the buffer length to 1 and the element
;; SIZE comes from the `$t` immediate. That immediate is what made this built-in
;; impossible before — a canon import registered by bare name has no type, and
;; guessing a width would move the wrong number of bytes into a peer's memory.
;;
;; `future.read@1` names type index 1 = i32 in the bootstrap table a bare core
;; module gets (a real component supplies its own type section).
;;
;; Reading a future that has not been written answers BLOCKED, not zero — the
;; distinction a guest needs to tell "not ready" from "ready, value 0".

(module
  (import "canon" "future.new" (func $future_new (result i64)))
  (import "canon" "future.read@1" (func $future_read (param i32 i32) (result i32)))
  (import "canon" "future.write@1" (func $future_write (param i32 i32) (result i32)))
  (memory 1)
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
        (local $packed i64) (local $rd i32) (local $wr i32)
    call $future_new
    local.set $packed
    local.get $packed
    i32.wrap_i64
    local.set $rd
    local.get $packed
    i64.const 32
    i64.shr_u
    i32.wrap_i64
    local.set $wr

    ;; unwritten: BLOCKED, not a zero value
    local.get $rd
    i32.const 64
    call $future_read
    i32.const 0xffffffff call $vybe_check_i32

    ;; put an i32 at 0 and hand it to the writable end
    i32.const 0
    i32.const 0x2a
    i32.store
    local.get $wr
    i32.const 0
    call $future_write
    ;; COMPLETED(0) | (1 << 4) — one element, always
    i32.const 0x10 call $vybe_check_i32

    ;; now the readable end delivers it into a different address
    local.get $rd
    i32.const 64
    call $future_read
    i32.const 0x10 call $vybe_check_i32

    i32.const 64
    i32.load
    i32.const 0x2a call $vybe_check_i32
  )
)
