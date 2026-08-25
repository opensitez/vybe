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
;; `$t` here is type index 1 = i32 in the bootstrap table. This component
;; declares NO type space of its own, so the bootstrap stands; a component that
;; declares one REPLACES those four entries and numbers from zero.
;;
;; Reading a future that has not been written answers BLOCKED, not zero — the
;; distinction a guest needs to tell "not ready" from "ready, value 0".
;;
;; ⛔ REWRITTEN AS A COMPONENT 2026-08-23, and not cosmetically. The unwritten
;; read below asserts BLOCKED, and `Binary.md:317` is
;;     0x16 t:<typeidx> opts:<opts> => (canon future.read t opts) 🔀
;; where only the ASYNC form may answer BLOCKED — the synchronous one SUSPENDS.
;; A bare core module cannot spell `opts`, so this was the SYNC read and the
;; test had been failing on a real deadlock since the sync suspension landed.
;; `async` is `canonopt` 0x06, an option on the row rather than a second
;; built-in, so declaring it needs a canon section and therefore a component.
;;
;; The WRITE is deliberately left synchronous: it completes immediately (a
;; future write always has somewhere to put its one element), so it has no
;; occasion to block and needs no `async`. Marking it too would hide that
;; asymmetry.

(component
  (canon future.new   1       (core func $fnew))
  (canon future.read  1 async (core func $fread))
  (canon future.write 1       (core func $fwrite))

  (core module $m
    (import "canon" "new"   (func $future_new (result i64)))
    (import "canon" "read"  (func $future_read (param i32 i32) (result i32)))
    (import "canon" "write" (func $future_write (param i32 i32) (result i32)))
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
      i32.const 0x2a call $vybe_check_i32))

  (core instance (instantiate $m
    (with "canon" (instance
      (export "new"   (func $fnew))
      (export "read"  (func $fread))
      (export "write" (func $fwrite))))))
)
