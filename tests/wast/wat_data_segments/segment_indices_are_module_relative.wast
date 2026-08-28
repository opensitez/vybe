;; vybe-test: wast/wat_data_segments/segment_indices_are_module_relative
;; vybe-test-mode: run
;;
;; A `.wast` script is ONE program, but each `(module …)` numbers its own
;; memories, tables, data segments and element segments from 0. The compiler
;; pushes every data segment onto one script-wide list, so a written dataidx
;; needs the same per-module BASE that memories and tables already had.
;; It did not have one: the second module's `memory.init 0` copied the FIRST
;; module's segment.
;;
;; A single-module file cannot show this — base 0 is the identity — which is
;; why the whole existing corpus passed. Every module below therefore declares
;; its own segments, and the ones in the earlier modules hold DIFFERENT bytes,
;; so reading the wrong segment produces a wrong answer rather than the same
;; one.
;;
;; The other half is the memidx: `memory.init x y` names a memory and a data
;; segment, `memory.init y` names only a data segment. Nothing but the COUNT of
;; bare indices tells them apart — read `memory.init $d` as a memidx and the op
;; loses its segment; ignore the memidx and `memory.init $mem2 $d` writes to
;; memory 0. Both spellings are below.

;; ── module 1: the decoy ──────────────────────────────────────────────────
;; Its passive segment is data index 0 script-wide. Later modules must NOT
;; reach it.
(module
  (memory 1)
  (data "\11\11\11\11")
  (func (export "m1_b") (param $at i32) (result i32) (i32.load8_u (local.get $at))))

;; ── module 2: numeric dataidx ────────────────────────────────────────────
(module
  (memory 1)
  (data "\aa\bb\cc\dd")
  (func (export "init") (param $dst i32) (param $src i32) (param $n i32)
    (memory.init 0 (local.get $dst) (local.get $src) (local.get $n)))
  (func (export "b") (param $at i32) (result i32) (i32.load8_u (local.get $at))))

;; Copy two bytes starting at source offset 1 — `\bb\cc`, not `\aa\bb`, so a
;; wrong source offset is visible too.
(invoke "init" (i32.const 0) (i32.const 1) (i32.const 2))
(assert_return (invoke "b" (i32.const 0)) (i32.const 0xbb))
(assert_return (invoke "b" (i32.const 1)) (i32.const 0xcc))
;; Reading module 1's segment would put 0x11 here.
(assert_return (invoke "b" (i32.const 2)) (i32.const 0))

;; ── module 3: named segments, named memories ─────────────────────────────
;; `(memory.init $mem2 $d2 …)` names BOTH. `$d1` and `$d2` hold different
;; bytes and `$mem1`/`$mem2` are different memories, so a mistake in either
;; immediate shows up as a different byte or a byte in the wrong memory.
(module
  (memory $mem1 1)
  (memory $mem2 1)
  (data $d1 "\01\02\03\04")
  (data $d2 "\05\06\07\08")

  (func (export "init_1") (memory.init $mem1 $d1 (i32.const 0) (i32.const 0) (i32.const 4)))
  (func (export "init_2") (memory.init $mem2 $d2 (i32.const 0) (i32.const 0) (i32.const 4)))
  ;; The one-immediate spelling: dataidx only, default memory (= $mem1).
  (func (export "init_default") (memory.init $d2 (i32.const 8) (i32.const 0) (i32.const 4)))

  (func (export "b1") (param $at i32) (result i32) (i32.load8_u $mem1 (local.get $at)))
  (func (export "b2") (param $at i32) (result i32) (i32.load8_u $mem2 (local.get $at))))

(invoke "init_1")
(invoke "init_2")
(assert_return (invoke "b1" (i32.const 0)) (i32.const 1))
(assert_return (invoke "b1" (i32.const 3)) (i32.const 4))
(assert_return (invoke "b2" (i32.const 0)) (i32.const 5))
(assert_return (invoke "b2" (i32.const 3)) (i32.const 8))
;; …and neither write reached the other memory.
(assert_return (invoke "b2" (i32.const 4)) (i32.const 0))
(assert_return (invoke "b1" (i32.const 4)) (i32.const 0))

;; The abbreviated form lands on the DEFAULT memory with the named segment.
(invoke "init_default")
(assert_return (invoke "b1" (i32.const 8)) (i32.const 5))
(assert_return (invoke "b1" (i32.const 11)) (i32.const 8))
(assert_return (invoke "b2" (i32.const 8)) (i32.const 0))

;; ── module 4: active segments occupy index slots too ─────────────────────
;; WASM gives active and passive data segments ONE index space. If the base
;; only counted passive segments, this module's `$p` would resolve one slot
;; low and copy module 4's own ACTIVE bytes instead.
(module
  (memory 1)
  (data (i32.const 64) "\ff\ff\ff\ff")   ;; active — still consumes an index
  (data $p "\21\22\23\24")
  (func (export "init") (memory.init $p (i32.const 0) (i32.const 0) (i32.const 4)))
  (func (export "b") (param $at i32) (result i32) (i32.load8_u (local.get $at))))

(assert_return (invoke "b" (i32.const 64)) (i32.const 0xff))
(invoke "init")
(assert_return (invoke "b" (i32.const 0)) (i32.const 0x21))
(assert_return (invoke "b" (i32.const 3)) (i32.const 0x24))

;; ── module 1 is still intact ─────────────────────────────────────────────
;; Nothing above should have written into an earlier module's memory.
(assert_return (invoke "m1_b" (i32.const 0)) (i32.const 0))
