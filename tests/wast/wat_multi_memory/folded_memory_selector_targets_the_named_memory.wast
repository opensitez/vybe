;; vybe-test: wast/wat_multi_memory/folded_memory_selector_targets_the_named_memory
;; vybe-test-mode: run
;;
;; A multi-memory selector must be peeled in the FOLDED spelling too.
;;
;; `i32.store 1 …` names memory 1. The walker turns that leading bare index
;; into the `@@mem1` name suffix the emitter reads (`peel_mem_selector`), and
;; it did so at the plain-instruction parse site only. `walk_folded_core` —
;; the S-expression path — walked its `instr_arg` children straight into the
;; argument list, so `(i32.store 1 (i32.const 0) (i32.const 42))` reached the
;; emitter with the SELECTOR AS AN OPERAND: the op emitted memidx 0, and the
;; extra argument was pushed and dropped.
;;
;; The failure is silent in the worst way. A write meant for memory 1 lands in
;; memory 0, corrupting it, and the read-back from memory 1 answers whatever
;; memory 1 was initialised with — usually a plausible zero, and, when memory 1
;; is the larger of the two, an out-of-bounds trap against memory 0's limit
;; instead.
;;
;; Our own corpus missed it because every multi-memory case here was written
;; PLAIN (`i32.const 0 i32.const 222 i32.store 1`), and the plain path was the
;; one that worked. The spec's files fold freely, which is how it surfaced.
;;
;; Every op that carries a memidx is exercised in both spellings, so a future
;; fix that repairs one path and drops the other cannot pass.

(module
  ;; Deliberately different sizes: a selector that silently falls back to
  ;; memory 0 shows up as a bounds trap as well as a wrong value.
  (memory $a 1)
  (memory $b 2)

  (data (memory $b) (i32.const 256) "\aa\bb\cc\dd")

  ;; ── Loads and stores, folded ────────────────────────────────────────
  (func (export "store_folded_b") (i32.store 1 (i32.const 0) (i32.const 42)))
  (func (export "load_folded_a") (result i32) (i32.load 0 (i32.const 0)))
  (func (export "load_folded_b") (result i32) (i32.load 1 (i32.const 0)))

  ;; The same, by NAME rather than by number — `$b` has to resolve to the
  ;; same index before it can be mangled.
  (func (export "store_folded_named") (i32.store $b (i32.const 8) (i32.const 77)))
  (func (export "load_folded_named") (result i32) (i32.load $b (i32.const 8)))
  (func (export "load_folded_named_a") (result i32) (i32.load $a (i32.const 8)))

  ;; Narrow accesses take the selector the same way.
  (func (export "store8_folded_b") (i32.store8 1 (i32.const 16) (i32.const 0xff)))
  (func (export "load8u_folded_b") (result i32) (i32.load8_u 1 (i32.const 16)))
  (func (export "load8u_folded_a") (result i32) (i32.load8_u 0 (i32.const 16)))

  ;; A selector must not be confused with an `offset=` memarg sharing the slot.
  (func (export "store_folded_b_offset")
    (i32.store 1 offset=4 (i32.const 32) (i32.const 99)))
  (func (export "load_folded_b_offset") (result i32)
    (i32.load 1 offset=4 (i32.const 32)))

  ;; ── memory.size / memory.grow, folded ───────────────────────────────
  ;; `$a` is 1 page and `$b` is 2, so a lost selector answers the wrong one.
  (func (export "size_folded_a") (result i32) (memory.size 0))
  (func (export "size_folded_b") (result i32) (memory.size 1))
  (func (export "grow_folded_b") (result i32) (memory.grow 1 (i32.const 1)))

  ;; ── memory.fill, folded ─────────────────────────────────────────────
  (func (export "fill_folded_b")
    (memory.fill 1 (i32.const 64) (i32.const 7) (i32.const 4)))
  (func (export "load8u_fill_b") (result i32) (i32.load8_u 1 (i32.const 64)))
  (func (export "load8u_fill_a") (result i32) (i32.load8_u 0 (i32.const 64)))

  ;; ── memory.copy, folded — dst then src ──────────────────────────────
  ;; Copies the data segment out of `$b` and into `$a`, which is the only
  ;; direction that tells the two selectors apart.
  (func (export "copy_folded_b_to_a")
    (memory.copy 0 1 (i32.const 128) (i32.const 256) (i32.const 4)))
  (func (export "load_copy_a") (result i32) (i32.load 0 (i32.const 128)))
  (func (export "load_copy_b_src") (result i32) (i32.load 1 (i32.const 256)))

  ;; ── Plain spelling, as the control ──────────────────────────────────
  ;; This path always worked. If a fix trades one spelling for the other,
  ;; these fail instead.
  (func (export "store_plain_b")
    i32.const 512
    i32.const 1234
    i32.store 1)
  (func (export "load_plain_b") (result i32)
    i32.const 512
    i32.load 1)
  (func (export "load_plain_a") (result i32)
    i32.const 512
    i32.load 0)
)

;; The data segment landed in `$b`, not `$a`.
(assert_return (invoke "load_copy_b_src") (i32.const 0xddccbbaa))

(invoke "store_folded_b")
(assert_return (invoke "load_folded_b") (i32.const 42))
(assert_return (invoke "load_folded_a") (i32.const 0))

(invoke "store_folded_named")
(assert_return (invoke "load_folded_named") (i32.const 77))
(assert_return (invoke "load_folded_named_a") (i32.const 0))

(invoke "store8_folded_b")
(assert_return (invoke "load8u_folded_b") (i32.const 0xff))
(assert_return (invoke "load8u_folded_a") (i32.const 0))

(invoke "store_folded_b_offset")
(assert_return (invoke "load_folded_b_offset") (i32.const 99))

(assert_return (invoke "size_folded_a") (i32.const 1))
(assert_return (invoke "size_folded_b") (i32.const 2))
(assert_return (invoke "grow_folded_b") (i32.const 2))
(assert_return (invoke "size_folded_b") (i32.const 3))
(assert_return (invoke "size_folded_a") (i32.const 1))

(invoke "fill_folded_b")
(assert_return (invoke "load8u_fill_b") (i32.const 7))
(assert_return (invoke "load8u_fill_a") (i32.const 0))

(invoke "copy_folded_b_to_a")
(assert_return (invoke "load_copy_a") (i32.const 0xddccbbaa))

(invoke "store_plain_b")
(assert_return (invoke "load_plain_b") (i32.const 1234))
(assert_return (invoke "load_plain_a") (i32.const 0))
