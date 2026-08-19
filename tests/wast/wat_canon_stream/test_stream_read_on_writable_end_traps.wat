;; vybe-test: wast/wat_canon_stream/test_stream_read_on_writable_end_traps
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §stream_copy — `trap_if(not isinstance(e, EndT))`
;;
;; Reading the WRITABLE end is a trap, not a quiet nothing. Before the handle
;; table was consulted at all, `stream.read` accepted only the high-level
;; object form and answered Null for anything else — so passing the wrong end,
;; or any unrecognised handle, read as EOF. A guest cannot distinguish that
;; from a finished stream.

(module
  (import "canon" "stream.new" (func $stream_new (result i64)))
  (import "canon" "stream.read" (func $stream_read (param i32 i32 i32) (result i32)))
  (memory 1)
  (func (export "_start")
        (local $packed i64) (local $wr i32)
    call $stream_new
    local.set $packed
    local.get $packed
    i64.const 32
    i64.shr_u
    i32.wrap_i64
    local.set $wr

    ;; the writable end is not readable
    local.get $wr
    i32.const 0
    i32.const 8
    call $stream_read
    drop
  )
)
