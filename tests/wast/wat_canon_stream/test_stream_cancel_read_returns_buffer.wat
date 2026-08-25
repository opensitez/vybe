;; vybe-test: wast/wat_canon_stream/test_stream_cancel_read_returns_buffer
;; hand-written (the extractor is deprecated) against
;; proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon stream.{read,write}                  — e.state = COPYING before the copy
;;   §canon {stream,future}.cancel-{read,write}  — $f : (func (param i32) (result i32))
;;
;; The COPY LIFECYCLE, which is a state machine on each END and not on the
;; stream:
;;
;;   read with nothing buffered  → BLOCKED, and the end is left COPYING
;;   cancel-read while COPYING   → CANCELLED, and the end returns to IDLE
;;   read again                  → legal, because the end is IDLE again
;;
;; Returning BLOCKED while staying IDLE (what this used to do) would quietly
;; permit both a second concurrent read AND a cancel with nothing in flight.
;; Cancel is also NOT close: both ends stay usable afterwards, which is the
;; whole point of the CANCELLED code — it tells wasm the buffer is its own
;; again.
;;
;; ⛔ REWRITTEN AS A COMPONENT 2026-08-23, and this is not cosmetic. The test
;; asserted BLOCKED from a bare `(module …)` — but `Binary.md:310` is
;;     0x0f t:<typeidx> opts:<opts> => (canon stream.read t opts) 🔀
;; and only the ASYNC form may answer BLOCKED; the synchronous one SUSPENDS.
;; A bare core module cannot spell `opts` at all, so the read here was the SYNC
;; one and the test had been failing on a real deadlock ever since the sync
;; suspension was wired. `async` is `canonopt` 0x06 — an option on the row, not
;; a second built-in — so the fix is to declare it, which needs a canon
;; section, which needs a component.
;;
;; The `with` clause names each row by its binder; there is no `@N` here.

(component
  (canon stream.new         1       (core func $snew))
  (canon stream.read        1 async (core func $sread))
  (canon stream.cancel-read 1       (core func $scancel))

  (core module $m
    (import "canon" "new"    (func $stream_new (result i64)))
    (import "canon" "read"   (func $stream_read (param i32 i32 i32) (result i32)))
    (import "canon" "cancel" (func $stream_cancel_read (param i32) (result i32)))
    (memory 1)

    (func $vybe_check_i32 (param i32) (param i32)
      local.get 0
      local.get 1
      i32.ne
      if
        unreachable
      end)

    (func (export "_start") (local $packed i64) (local $rd i32)
      call $stream_new
      local.set $packed
      local.get $packed
      i32.wrap_i64
      local.set $rd

      ;; nothing has been written, and the writable end is still open, so the
      ;; copy cannot complete synchronously
      local.get $rd
      i32.const 0
      i32.const 8
      call $stream_read
      i32.const 0xffffffff call $vybe_check_i32

      ;; the copy is in flight, so it can be cancelled: CANCELLED(2), 0 copied
      local.get $rd
      call $stream_cancel_read
      i32.const 2 call $vybe_check_i32

      ;; and the end is IDLE again, so a second read is legal rather than a trap
      local.get $rd
      i32.const 0
      i32.const 8
      call $stream_read
      i32.const 0xffffffff call $vybe_check_i32))

  (core instance (instantiate $m
    (with "canon" (instance
      (export "new"    (func $snew))
      (export "read"   (func $sread))
      (export "cancel" (func $scancel))))))
)
