;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_mismatching_end_label
;; vybe-test-mode: run
;;
;; In the plain (non-folded) form a block may be CLOSED with its own label:
;; `block $l … end $l`. The id after `end` — and after `else` — must be the id
;; the block was opened with. The spec calls a disagreement "mismatching
;; label" and classifies it MALFORMED, a property of the text, not invalid.
;; An UNNAMED block mismatches every id, because it has none to match.
;;
;; The grammar cannot express this: `end` is an instruction name and the id
;; after it is an instruction argument, so nothing relates the two. It is
;; checked on the token stream instead, which is also why only the plain form
;; can be wrong — a folded `(block $l …)` is closed by its paren and never
;; writes a closing id.
;;
;; ⛔ The dangerous direction here is over-flagging: a malformed check that
;; rejects too much makes the assertion pass for the WRONG reason and hides
;; whatever the fixture was really about. Every rejection below is therefore
;; paired with a CONTROL — the same shape, well-formed — that must still
;; compile and run.

;; ── the controls: well-formed, and they must stay that way ─────────────
(module
  ;; A block closed with no id at all.
  (func (export "plain") (result i32)
    block (result i32) i32.const 1 end)
  ;; A block closed with its OWN id.
  (func (export "matched") (result i32)
    block $a (result i32) i32.const 2 end $a)
  ;; Nested, each closing with its own — the inner one first.
  (func (export "nested") (result i32)
    block $outer (result i32)
      block $inner (result i32) i32.const 3 end $inner
    end $outer)
  ;; A named block whose closing id is omitted is fine — the id is optional.
  (func (export "named_open_bare_close") (result i32)
    block $b (result i32) i32.const 4 end)
  ;; loop and if, both spellings of the closing id.
  (func (export "loop_matched") (result i32)
    loop $l (result i32) i32.const 5 end $l)
  (func (export "if_matched") (result i32)
    i32.const 1 if $c (result i32) i32.const 6 else $c i32.const 7 end $c)
  ;; The folded form never writes a closing id and must not be touched.
  (func (export "folded") (result i32)
    (block $f (result i32) (i32.const 8)))
)
(assert_return (invoke "plain") (i32.const 1))
(assert_return (invoke "matched") (i32.const 2))
(assert_return (invoke "nested") (i32.const 3))
(assert_return (invoke "named_open_bare_close") (i32.const 4))
(assert_return (invoke "loop_matched") (i32.const 5))
(assert_return (invoke "if_matched") (i32.const 6))
(assert_return (invoke "folded") (i32.const 8))

;; ── the rejections ─────────────────────────────────────────────────────
;; An UNNAMED block closed with an id.
(assert_malformed (module quote "(func block end $l)") "mismatching label")
(assert_malformed (module quote "(func loop end $l)") "mismatching label")
(assert_malformed (module quote "(func i32.const 0 if end $l)") "mismatching label")

;; A NAMED block closed with a different id.
(assert_malformed (module quote "(func block $a end $l)") "mismatching label")
(assert_malformed (module quote "(func loop $a end $l)") "mismatching label")

;; `else` names the same label as its `if`.
(assert_malformed
  (module quote "(func i32.const 0 if $a else $b end $a)")
  "mismatching label"
)

;; The mismatch is with the INNERMOST open block, not with any open block:
;; `$outer` is open here, but the `end` being written closes `$inner`.
(assert_malformed
  (module quote "(func block $outer block $inner end $outer end $outer)")
  "mismatching label"
)

;; …and closing an unnamed inner block with the OUTER block's id is the same
;; mistake, which is the case a check that only looked at the outermost frame
;; would let through.
(assert_malformed
  (module quote "(func block $outer block end $outer end $outer)")
  "mismatching label"
)
