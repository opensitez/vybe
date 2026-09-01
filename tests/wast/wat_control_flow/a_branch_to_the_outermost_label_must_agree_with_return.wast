;; vybe-test: wast/wat_control_flow/a_branch_to_the_outermost_label_must_agree_with_return
;; vybe-test-mode: run
;;
;; ⛔ A BRANCH TO THE FUNCTION'S OWN LABEL IS A `return`, NOT A FALL-THROUGH.
;;
;; §4.4.8: `br l` where `l` names the outermost label — depth `len(labels)`,
;; the function's implicit block — exits the function with its results, exactly
;; as `return` does. Every branch emitter here resolved a target through
;; `LabelStack::resolve`, which answers `None` for BOTH "out of scope" and
;; "the function's own label", and then lowered that single `None` to
;; `Break(BreakTarget::Implicit)`. At function top level that breaks nothing:
;; `(br 0 (i32.const 50)) (i32.const 51)` returned 51.
;;
;; Nothing upstream could catch it. The module is VALID, so the stack-typing
;; pass is right to accept it; the failure is a WRONG VALUE. And `return` was
;; correct all along, as was every branch to a `block` or `loop` — only the
;; sibling spelling was wrong, which is the same bug class as `(elem 0)` vs
;; `(elem $a)`: one spelling carries the logic, the other silently does not.
;;
;; So this pins the AGREEMENT rather than either side. It fails if a branch to
;; the outermost label ever diverges from `return` again — in EITHER
;; direction, including a "fix" that breaks `return` to match a broken `br`.

(module
  ;; ── single result ──────────────────────────────────────────────────────
  (func (export "by_return") (result i32)
    (return (i32.const 50)) (i32.const 51)
  )
  (func (export "by_br") (result i32)
    (br 0 (i32.const 50)) (i32.const 51)
  )
  (func (export "by_br_if") (param i32) (result i32)
    (drop (br_if 0 (i32.const 50) (local.get 0))) (i32.const 51)
  )
  (func (export "by_br_table") (result i32)
    (drop (br_table 0 (i32.const 50) (i32.const 0))) (i32.const 51)
  )
  ;; Depth counts the enclosing frames: 1 from inside the loop is the function.
  (func (export "by_br_table_from_loop") (result i32)
    (loop (result i32) (br_table 1 1 (i32.const 50) (i32.const 0)) (i32.const 51))
  )

  ;; ── multi-value results ────────────────────────────────────────────────
  (func (export "multi_by_return") (result i32 i64)
    (return (i32.const 50) (i64.const 51)) (i32.const 60) (i64.const 61)
  )
  (func (export "multi_by_br_if") (param i32) (result i32 i64)
    (drop (drop (br_if 0 (i32.const 50) (i64.const 51) (local.get 0))))
    (i32.const 60) (i64.const 61)
  )

  ;; ── no result: the branch must still UNWIND ────────────────────────────
  ;; A value-carrying fix alone passes everything above and still lets this
  ;; one fall through into the store. `n == 0` is a bare `return`, which
  ;; discards the 3 and the 1i64 the branch unwinds past.
  (global $g (mut i32) (i32.const 0))
  (func (export "void_by_br_table")
    (i32.const 3) (i64.const 1) (br_table 0 (i32.const 0))
    (global.set $g (i32.const 7))
  )
  (func (export "g") (result i32) (global.get $g))

  ;; ── the REST of the branch family, which was left behind once ──────────
  ;; ⛔ `br`/`br_if`/`br_table` were fixed and `br_on_null`/`br_on_non_null`
  ;; were NOT, in the same file, by the same edit. Nothing caught it: the GC
  ;; `br_on_cast` fixtures stop on an earlier validation gap and
  ;; `br_on_non_null.wast` fails on a separate bug of its own, so the whole
  ;; family looked covered while half of it still fell through. Every branch
  ;; instruction that can name the outermost label is listed here for that
  ;; reason — the omission, not the bug, is what this section guards.
  (type $t (func (param i32) (result i32)))
  (func $sq (type $t) (i32.mul (local.get 0) (local.get 0)))
  (elem declare func $sq)

  (func (export "on_null_taken") (result i32)
    (drop (br_on_null 0 (i32.const 50) (ref.null func))) (i32.const 51)
  )
  ;; br_on_non_null delivers `t*` AND the reference, so the outermost label
  ;; takes two results — returning the ref alone is right only for a
  ;; single-result function.
  (func $on_non_null (param $r (ref null $t)) (result i32 (ref $t))
    (br_on_non_null 0 (i32.const 50) (local.get $r))
    (i32.const 51) (ref.func $sq)
  )
  (func (export "on_non_null_taken") (result i32)
    (call_ref $t (call $on_non_null (ref.func $sq)))
  )
  (func (export "on_non_null_fell") (result i32)
    (call_ref $t (call $on_non_null (ref.null $t)))
  )

  ;; ── the untaken path must be untouched ─────────────────────────────────
  ;; `br_if` PEEKS the carried values; consuming them would corrupt the stack
  ;; the fall-through still owns.
  (func (export "not_taken") (param i32) (result i32)
    (drop (br_if 0 (i32.const 50) (local.get 0))) (i32.const 51)
  )
)

;; Every spelling agrees with `return`.
(assert_return (invoke "by_return") (i32.const 50))
(assert_return (invoke "by_br") (i32.const 50))
(assert_return (invoke "by_br_if" (i32.const 1)) (i32.const 50))
(assert_return (invoke "by_br_table") (i32.const 50))
(assert_return (invoke "by_br_table_from_loop") (i32.const 50))

(assert_return (invoke "multi_by_return") (i32.const 50) (i64.const 51))
(assert_return (invoke "multi_by_br_if" (i32.const 1)) (i32.const 50) (i64.const 51))

;; The rest of the branch family agrees too.
(assert_return (invoke "on_null_taken") (i32.const 50))
(assert_return (invoke "on_non_null_taken") (i32.const 2500))
(assert_return (invoke "on_non_null_fell") (i32.const 2601))

;; The unwind happened, so the store below the branch never ran.
(assert_return (invoke "void_by_br_table"))
(assert_return (invoke "g") (i32.const 0))

;; Untaken falls through, and the value it did not carry is still intact.
(assert_return (invoke "not_taken" (i32.const 0)) (i32.const 51))
(assert_return (invoke "multi_by_br_if" (i32.const 0)) (i32.const 60) (i64.const 61))
