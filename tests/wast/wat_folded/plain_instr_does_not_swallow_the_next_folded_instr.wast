;; vybe-test: wast/wat_folded/plain_instr_does_not_swallow_the_next_folded_instr
;; vybe-test-mode: run
;;
;; A PLAIN instruction takes no folded operands. The spec's text format folds
;; only into an explicit `( … )`: written in a sequence, `nop (i32.const 5)` is
;; two instructions — `nop`, then `i32.const 5` — not a `nop` applied to 5.
;;
;; Our grammar spells `plain_instr = instr_name ~ instr_arg*` with `folded_instr`
;; among the `instr_arg` alternatives, so the following folded instruction is
;; parsed INTO the plain one. For an instruction that pops nothing that is a
;; silent deletion: the whole `(i32.const 5)` disappeared and the enclosing
;; `(result i32)` block came out empty, returning null. The bytecode was
;; literally `block / end` — no trace that anything was dropped.
;;
;; Nothing caught this because our own corpus writes sequences in one style or
;; the other: all-folded (`(nop) (i32.const 5)`) and all-plain
;; (`nop i32.const 5`) both work. Only the MIX exposes it, and the mix is what
;; the spec's own files use freely.
;;
;; Spec-format so `wasmtime wast` arbitrates every case.

(module
  (global $g (mut i32) (i32.const 0))

  ;; ── The minimal case: everything after the `nop` vanished ────────────
  (func (export "nop_then_folded") (result i32)
    (block (result i32)
      nop
      (i32.const 5)))

  ;; More than one following instruction, and one that consumes the others.
  (func (export "nop_then_several_folded") (result i32)
    (block (result i32)
      nop
      (i32.const 1)
      (i32.const 4)
      i32.add))

  (func (export "two_nops_then_folded") (result i32)
    (block (result i32)
      nop
      nop
      (i32.const 5)))

  ;; ── `nop` must not consume from the stack either ─────────────────────
  ;; It was falling to the arity default of 1, so it drained the pending
  ;; value before discarding it.
  (func (export "folded_then_nop") (result i32)
    (block (result i32)
      (i32.const 5)
      nop))

  (func (export "nop_between_operands") (result i32)
    (block (result i32)
      (i32.const 2)
      nop
      (i32.const 3)
      i32.add))

  (func (export "nop_does_not_disturb_a_deeper_stack") (result i32)
    (block (result i32)
      (i32.const 10)
      (i32.const 20)
      nop
      i32.sub))

  ;; ── The two spellings that always worked, as controls ────────────────
  ;; If a fix regressed either of these it would be trading one silent drop
  ;; for another.
  (func (export "all_folded") (result i32)
    (block (result i32)
      (nop)
      (i32.const 5)))

  (func (export "all_plain") (result i32)
    (block (result i32)
      nop
      i32.const 5))

  ;; ── The same mix outside a block, at function level ──────────────────
  (func (export "nop_then_folded_at_function_level") (result i32)
    nop
    (i32.const 5))

  ;; ── A `nop` must not swallow a folded instruction with side effects ───
  ;; The deletion is worse than a wrong value here: the store never happens.
  (func (export "nop_then_effect")
    nop
    (global.set $g (i32.const 42)))
  (func (export "g") (result i32) (global.get $g))
)

(assert_return (invoke "nop_then_folded") (i32.const 5))
(assert_return (invoke "nop_then_several_folded") (i32.const 5))
(assert_return (invoke "two_nops_then_folded") (i32.const 5))
(assert_return (invoke "folded_then_nop") (i32.const 5))
(assert_return (invoke "nop_between_operands") (i32.const 5))
(assert_return (invoke "nop_does_not_disturb_a_deeper_stack") (i32.const -10))
(assert_return (invoke "all_folded") (i32.const 5))
(assert_return (invoke "all_plain") (i32.const 5))
(assert_return (invoke "nop_then_folded_at_function_level") (i32.const 5))

(assert_return (invoke "g") (i32.const 0))
(invoke "nop_then_effect")
(assert_return (invoke "g") (i32.const 42))
