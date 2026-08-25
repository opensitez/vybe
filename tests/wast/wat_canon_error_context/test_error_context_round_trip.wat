;; vybe-test: wast/wat_canon_error_context/test_error_context_round_trip
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §📝 canon error-context.new / .debug-message / .drop (:5147, :5189, :5215)
;;
;; The three 📝 rows, round-tripped: a message goes in as a (ptr, len) range,
;; comes back out as a (ptr, len) pair written into memory, and the handle is
;; released.
;;
;; ⛔ WHAT THIS TEST DELIBERATELY PINS: the message SURVIVES.
;;
;; `canon_error_context_new` is explicitly nondeterministic —
;;
;;     if DETERMINISTIC_PROFILE or random.randint(0,1):
;;       s = String(('', 'utf8', 0))          # discard it
;;     else:
;;       s = host_defined_transformation(load_string_from_range(...))
;;
;; — so a host that always returned an EMPTY message would be conformant, and
;; every assert below would still pass if it only checked that the calls did
;; not trap. We preserve the message instead, which is equally conformant and
;; is the only choice that makes the feature worth having: an error-context
;; whose debug message is always empty aids no debugging.
;;
;; So the assert reads the message BACK — the length AND the first byte — and
;; both bite: expecting 13 instead of 12, or `e` instead of `d`, fails.
;;
;; Not the deterministic profile either. That flag also gates NaN scrambling,
;; so claiming it here would assert a whole execution profile this VM does not
;; implement.

(module
  (import "canon" "error-context.new"
    (func $ecnew (param i32 i32) (result i32)))
  (import "canon" "error-context.debug-message"
    (func $ecmsg (param i32 i32)))
  (import "canon" "error-context.drop"
    (func $ecdrop (param i32)))

  (memory (export "mem") 1)
  (data (i32.const 100) "disk on fire")

  ;; Build a context from the 12 bytes at 100, write its (ptr,len) pair to 200,
  ;; then release the handle.
  (func $run
    (local $h i32)
    (local.set $h (call $ecnew (i32.const 100) (i32.const 12)))
    (call $ecmsg (local.get $h) (i32.const 200))
    (call $ecdrop (local.get $h)))

  ;; The LENGTH half of the returned pair.
  (func (export "msg_len") (result i32)
    (call $run)
    (i32.load (i32.const 204)))

  ;; The FIRST BYTE of the returned message — `d` of "disk", 0x64.
  ;; Reads through the returned POINTER, so a host that wrote a pair pointing
  ;; at the wrong place fails here rather than passing on length alone.
  (func (export "msg_first_byte") (result i32)
    (call $run)
    (i32.load8_u (i32.load (i32.const 200))))
)

(assert_return (invoke "msg_len") (i32.const 12))
(assert_return (invoke "msg_first_byte") (i32.const 100))
