;; vybe-test: wast/wat_multi_memory/an_offset_and_a_non_default_memory_compose
;; vybe-test-mode: run
;;
;; A load/store can name BOTH a non-default memory and an `offset=`. Neither is
;; an AST argument — every `OperandFormat` arm in the emitter indexes its
;; immediates positionally, so an extra argument would silently shift a lane
;; index or a typeidx — so both ride across as name suffixes and both end up in
;; the same marker-tagged memarg: `i32.load@@off4@@mem1`.
;;
;; ⛔ TWO SUFFIXES NEED ONE READER. The emitter had two independent decoders of
;; this name — a routing site taking everything before `@@mem` as the opcode,
;; and a selector site splitting ON `@@mem` and parsing each remainder as a
;; number. With the offset written last the selector site parsed `"1@@off4"`,
;; failed, and dropped the memidx (so the access silently went to memory 0);
;; with it written first the routing site kept `@@off4` in the name and failed
;; to recognise the opcode at all. Whichever order, ONE of the two broke. They
;; share a decomposition now, and the base name is everything before the first
;; `@@` everywhere.

(module
  (memory $a 1)
  (memory $b 1)
  (data (memory $a) (i32.const 0) "AAAAAAAAAAAAAAAA")
  (data (memory $b) (i32.const 0) "bbbbbbbbbbbbbbbb")

  ;; offset alone, default memory — the control.
  (func (export "def_off") (param $i i32) (result i32)
    (i32.load8_u offset=3 (local.get $i)))
  ;; memory alone — the other control.
  (func (export "mem_only") (param $i i32) (result i32)
    (i32.load8_u $b (local.get $i)))
  ;; BOTH, folded and plain.
  (func (export "both") (param $i i32) (result i32)
    (i32.load8_u $b offset=3 (local.get $i)))
  (func (export "both_plain") (param $i i32) (result i32)
    local.get $i
    i32.load8_u $b offset=3)

  ;; A store into the non-default memory at an offset, read back through a
  ;; DIFFERENT instruction, so a shared mistake cannot cancel out.
  (func (export "store_both") (param $i i32) (param $v i32)
    (i32.store8 $b offset=3 (local.get $i) (local.get $v)))
  (func (export "read_b") (param $i i32) (result i32)
    (i32.load8_u $b (local.get $i)))
  (func (export "read_a") (param $i i32) (result i32)
    (i32.load8_u $a (local.get $i)))
)

;; Each memory holds a different byte, so a lost memidx is visible as a value
;; and not only as a bounds trap.
(assert_return (invoke "def_off" (i32.const 0)) (i32.const 65))   ;; 'A', memory $a
(assert_return (invoke "mem_only" (i32.const 0)) (i32.const 98))  ;; 'b', memory $b
(assert_return (invoke "both" (i32.const 0)) (i32.const 98))
(assert_return (invoke "both_plain" (i32.const 0)) (i32.const 98))

;; The offset displaces within the memory it named: byte 3+3 = 6 is still 'b',
;; byte 3+13 = 16 is past the data and reads 0.
(assert_return (invoke "both" (i32.const 3)) (i32.const 98))
(assert_return (invoke "both" (i32.const 13)) (i32.const 0))

;; The write lands in $b at 5+3 = 8 — and NOT in $a, which a dropped memidx
;; would have hit.
(invoke "store_both" (i32.const 5) (i32.const 90))
(assert_return (invoke "read_b" (i32.const 8)) (i32.const 90))
(assert_return (invoke "read_a" (i32.const 8)) (i32.const 65))
;; …and not at 5 either, which a dropped OFFSET would have hit.
(assert_return (invoke "read_b" (i32.const 5)) (i32.const 98))

;; The unsigned effective address applies per memory, not only to memory 0.
(assert_trap (invoke "both" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "store_both" (i32.const -1) (i32.const 1)) "out of bounds memory access")
