;; vybe-test: wast/wat_canon_stream/test_async_stream_read_answers_blocked
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon stream.{read,write}, and Binary.md:310
;;     0x0f t:<typeidx> opts:<opts> => (canon stream.read t opts (core func)) 🔀
;;
;; ⛔ `async` IS `canonopt` 0x06 — AN OPTION ON THE ROW, NOT A SECOND BUILT-IN.
;; There is exactly one `CanonBuiltin::StreamRead`; a second one would put the
;; same operation in two canonidx spaces. The distinction rides in `opts`, and
;; it could not arrive at all until the canon section had a producer.
;;
;; Only the ASYNC form may answer BLOCKED (0xffffffff). The synchronous form
;; must SUSPEND, precisely so it can return a real payload rather than a
;; sentinel the caller has to special-case.
;;
;; The discriminator is ONE TOKEN. Its control,
;; test_sync_stream_read_suspends_instead, is this file with `async` deleted
;; from the canon row and nothing else changed; it deadlocks instead of
;; answering. Neither test alone shows anything — a runtime that ignored `opts`
;; entirely would pass whichever one matched its single hardcoded behaviour.
;;
;; ⚠ The end stays COPYING after a BLOCKED: the copy IS in flight. A caller
;; wanting POSIX `EAGAIN` retry semantics must issue `stream.cancel-read`
;; before reading again — that is the only thing that returns an end to IDLE,
;; and a bare retry traps on "not IDLE".

(component
  (canon stream.new 1 (core func $snew))
  (canon stream.read 1 async (core func $sread))

  (core module $m
    (import "canon" "new"  (func $stream_new (result i64)))
    (import "canon" "read" (func $stream_read (param i32 i32 i32) (result i32)))
    (memory 1)
    (func (export "_start") (local $rd i32)
      (local.set $rd (i32.wrap_i64 (call $stream_new)))
      ;; Nothing written and the writable end still open, so the copy cannot
      ;; complete synchronously. Async ⇒ BLOCKED rather than a suspend.
      (if (i32.ne
            (call $stream_read (local.get $rd) (i32.const 0) (i32.const 8))
            (i32.const 0xffffffff))
        (then unreachable))))

  (core instance (instantiate $m
    (with "canon" (instance
      (export "new"  (func $snew))
      (export "read" (func $sread))))))
)
