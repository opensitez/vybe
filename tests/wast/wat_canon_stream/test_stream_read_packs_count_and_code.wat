;; vybe-test: wast/wat_canon_stream/test_stream_read_packs_count_and_code
;; hand-written (the extractor is deprecated) against
;; proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon {stream,future}.new  —  $f : (func (result i64)), [ ri | (wi << 32) ]
;;   §canon stream.{read,write}  —  $f : (func (param i32 T T) (result T))
;;
;; THE encoding test. `packed_result = result | (buffer.progress << 4)`, so a
;; one-element COMPLETED read is 0x10 — NOT 1. Returning the bare count is the
;; mistake that would make every module we emit disagree with a conforming
;; runtime on the very first element copied, invisibly, because both ends of
;; our own tests would share the error.
;;
;; It also pins `stream.new`'s shape: ONE i64 with the readable end in the low
;; 32 bits and the writable end in the high 32. Two bare i32 pushes — what this
;; used to do — is not a signature any conforming module can call.

(module
  (import "canon" "stream.new" (func $stream_new (result i64)))
  (import "canon" "stream.write" (func $stream_write (param i32 i32 i32) (result i32)))
  (import "canon" "stream.read" (func $stream_read (param i32 i32 i32) (result i32)))
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

    ;; readable end = low 32 bits
    local.get $packed
    i32.wrap_i64
    local.set $rd

    ;; writable end = high 32 bits
    local.get $packed
    i64.const 32
    i64.shr_u
    i32.wrap_i64
    local.set $wr

    ;; the two ends are distinct handles
    local.get $rd
    local.get $wr
    i32.eq
    if
      unreachable
    end

    ;; one element in, FROM LINEAR MEMORY — `stream.write` is (handle, ptr, n),
    ;; the mirror of `stream.read`, and the spec runs both through one
    ;; `stream_copy`. Handing the element over as a value is not a signature any
    ;; conforming component could satisfy.
    i32.const 128
    i32.const 65
    i32.store8
    local.get $wr
    i32.const 128
    i32.const 1
    call $stream_write
    ;; the write reports its own packed result: COMPLETED(0) | (1 << 4)
    i32.const 0x10 call $vybe_check_i32

    ;; ...and back out into linear memory at 0
    local.get $rd
    i32.const 0
    i32.const 8
    call $stream_read
    i32.const 0x10 call $vybe_check_i32

    ;; the byte really landed
    i32.const 0
    i32.load8_u
    i32.const 65 call $vybe_check_i32
  )
)
