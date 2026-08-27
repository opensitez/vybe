;; vybe-test: wast/wat_folded/folded_branches_must_reach_the_statement_lowering
;; vybe-test-mode: run
;;
;; A folded instruction is walked as an EXPRESSION unless its head is on the
;; list in `folded_needs_stmt_lowering`. Anything that BRANCHES has to be on
;; that list: a branch has no expression form, and the walk turns it into an
;; ordinary call whose value the enclosing block then adopts as its result.
;;
;; `br_on_cast` / `br_on_cast_fail` were missing. The damage is not "the
;; branch does nothing" — it is worse. The block's result assignment is
;; appended AFTER the body, so
;;
;;     (block $l (result (ref i31))
;;       (br_on_cast $l anyref (ref i31) X)
;;       (return (i32.const -1)))
;;
;; lowered to `return -1` FOLLOWED BY `result = br_on_cast(...)`: the
;; instruction the branch was supposed to jump over ran first, unconditionally.
;; `gc/br_on_cast.wast` writes every case folded, so the whole file answered
;; the fall-through value.
;;
;; Both spellings are exercised for each op — the flat one already worked, and
;; a fix that traded one for the other would pass a folded-only test.

(module
  (type $st (struct (field i32)))

  ;; ── br_on_cast, folded ───────────────────────────────────────────
  (func (export "i31_folded") (param $take i32) (result i32)
    (block $l (result (ref i31))
      (br_on_cast $l anyref (ref i31)
        (if (result anyref) (local.get $take)
          (then (ref.i31 (i32.const 7)))
          (else (struct.new $st (i32.const 3)))))
      (return (i32.const -1)))
    (i31.get_u))

  ;; ── the same, written flat ───────────────────────────────────────
  ;; This spelling always worked; it is here so a fix that trades one parse
  ;; site for the other cannot pass.
  (func (export "i31_flat") (param $take i32) (result i32)
    (block $l (result (ref i31))
      local.get $take
      if (result anyref)
        i32.const 7
        ref.i31
      else
        i32.const 3
        struct.new $st
      end
      br_on_cast $l anyref (ref i31)
      drop
      i32.const -1
      return
    )
    (i31.get_u))

  ;; ── br_on_cast_fail, folded: branches when the cast FAILS ────────
  (func (export "fail_folded") (param $take i32) (result i32)
    (block $l (result anyref)
      (drop
        (br_on_cast_fail $l anyref (ref i31)
          (if (result anyref) (local.get $take)
            (then (ref.i31 (i32.const 7)))
            (else (struct.new $st (i32.const 3))))))
      ;; The cast SUCCEEDED — fall through without branching.
      (return (i32.const 100)))
    ;; Reached only via the branch, i.e. the cast failed.
    (drop)
    (i32.const 200))
)

(assert_return (invoke "i31_folded" (i32.const 1)) (i32.const 7))
(assert_return (invoke "i31_folded" (i32.const 0)) (i32.const -1))
(assert_return (invoke "i31_flat" (i32.const 1)) (i32.const 7))
(assert_return (invoke "i31_flat" (i32.const 0)) (i32.const -1))
(assert_return (invoke "fail_folded" (i32.const 1)) (i32.const 100))
(assert_return (invoke "fail_folded" (i32.const 0)) (i32.const 200))
