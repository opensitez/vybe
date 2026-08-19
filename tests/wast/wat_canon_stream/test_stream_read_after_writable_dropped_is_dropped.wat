;; vybe-test: wast/wat_canon_stream/test_stream_read_after_writable_dropped_is_dropped
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon stream.{read,write} — `if result == CopyResult.DROPPED: e.state = DONE`
;;
;; EOF is not a value, it is a CopyResult. Dropping the writable end signals
;; that no further copies are possible, so a read answers DROPPED(1) with a
;; count of 0 — and the readable end goes to DONE, after which anything but
;; `drop-*` traps.
;;
;; The old shape pushed a bare Null here, which a guest could not tell apart
;; from "read a null element" — and, worse, an i32 handle it did not recognise
;; produced exactly the same Null. Silent EOF for a perfectly live stream.

(module
  (import "canon" "stream.new" (func $stream_new (result i64)))
  (import "canon" "stream.read" (func $stream_read (param i32 i32 i32) (result i32)))
  (import "canon" "stream.drop-writable" (func $drop_wr (param i32)))
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
    call $stream_new
    local.set $packed
    local.get $packed
    i32.wrap_i64
    local.set $rd
    local.get $packed
    i64.const 32
    i64.shr_u
    i32.wrap_i64
    local.set $wr

    ;; close the write end — the reader can never receive anything now
    local.get $wr
    call $drop_wr

    ;; DROPPED(1) with progress 0
    local.get $rd
    i32.const 0
    i32.const 8
    call $stream_read
    i32.const 1 call $vybe_check_i32
  )
)
