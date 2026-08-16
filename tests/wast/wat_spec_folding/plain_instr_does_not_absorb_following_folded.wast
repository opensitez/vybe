;; vybe-test: wast/wat_spec_folding/plain_instr_does_not_absorb_following_folded
;; vybe-test-mode: run
;;
;; KNOWN-FAILING for `drop`. From `core/text/instructions.rst`: folding is an
;; abbreviation that applies only inside an explicit `( … )`. A PLAIN
;; instruction takes no folded operands, so in a sequence
;;
;;     drop (i32.const 5)
;;
;; is TWO instructions, not one. The grammar here spells
;; `plain_instr = instr_name ~ instr_arg*` with `folded_instr` among the
;; `instr_arg` alternatives, so the following folded instruction is parsed
;; INTO the plain one and then consumed as its operand.
;;
;; This is invisible in a corpus written all-folded or all-plain — both styles
;; give the right answer on their own. Only the MIX exposes it, and the mix is
;; what the spec's own text uses freely.
;;
;; `nop` was fixed earlier (it discarded what it absorbed, deleting the
;; instruction outright). `drop` is the remaining case and needs the grammar
;; change: `folded_instr` out of `instr_arg`, which first needs an explicit
;; `instr_arg` alternative for the reftype immediate `(ref null? ht)` — that
;; form currently matches only via `folded_instr`, so removing it without the
;; replacement stops `ref.test (ref $t)` from parsing.

(module
  ;; drop: `7 8 drop 5 add` → 7 8 → drop → 7 → 5 → 7 5 → add → 12
  (func (export "drop_then_folded") (result i32)
    (i32.const 7)
    (i32.const 8)
    drop
    (i32.const 5)
    i32.add)

  ;; The same computation written all-plain and all-folded, which already work.
  (func (export "all_plain") (result i32)
    i32.const 7
    i32.const 8
    drop
    i32.const 5
    i32.add)
  (func (export "all_folded") (result i32)
    (i32.add (i32.const 7) (i32.const 5)))

  ;; nop, fixed earlier — a plain `nop` must not delete the folded instruction
  ;; that follows it.
  (func (export "nop_then_folded") (result i32)
    nop
    (i32.const 5)
    (i32.const 7)
    i32.add)

  ;; local.set is void and takes one operand; the folded instruction after it
  ;; is a separate instruction supplying the RESULT.
  (func (export "localset_then_folded") (result i32) (local $x i32)
    (i32.const 3)
    local.set $x
    (i32.const 9)
    (local.get $x)
    i32.add)

  ;; A store is void with two operands, both already on the stack.
  (memory 1)
  (func (export "store_then_folded") (result i32)
    (i32.const 0)
    (i32.const 42)
    i32.store
    (i32.const 0)
    i32.load)

  ;; drop appearing twice in one sequence.
  (func (export "two_drops") (result i32)
    (i32.const 1)
    (i32.const 2)
    drop
    (i32.const 3)
    drop
    (i32.const 4)
    i32.add)
)

;; The mixed form must agree with both unmixed forms.
(assert_return (invoke "drop_then_folded") (i32.const 12))
(assert_return (invoke "all_plain") (i32.const 12))
(assert_return (invoke "all_folded") (i32.const 12))
(assert_return (invoke "nop_then_folded") (i32.const 12))
(assert_return (invoke "localset_then_folded") (i32.const 12))
(assert_return (invoke "store_then_folded") (i32.const 42))
;; 1 2 drop → 1 ; 3 drop → 1 ; 4 → 1 4 → add → 5
(assert_return (invoke "two_drops") (i32.const 5))
