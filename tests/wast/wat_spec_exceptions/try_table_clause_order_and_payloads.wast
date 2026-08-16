;; vybe-test: wast/wat_spec_exceptions/try_table_clause_order_and_payloads
;; vybe-test-mode: run
;;
;; From the exception-handling instructions in core/exec/instructions.rst.
;;
;; `try_table` is a BLOCK that installs one handler per CLAUSE for the extent
;; of its body. The rules that carry weight:
;;
;; * Matching is by TAG IDENTITY, and clauses are tried in the order written.
;;   Two tags with the same signature are still different tags.
;; * `catch` transfers the exception's payload to the label; `catch_ref` also
;;   pushes an `exnref`; `catch_all` pushes nothing; `catch_all_ref` pushes
;;   only the `exnref`.
;; * A handler is installed only for the DYNAMIC EXTENT of the body. Leaving
;;   the body — by falling off its `end` OR by a branch out of it — ends that
;;   extent, and a later throw must not be caught by a region already left.
;; * `throw_ref` re-raises a captured exception with its tag and payload
;;   intact, so a handler further out sees the ORIGINAL tag, not a new one.
;;
;; The branch-out case is the one that hides: a `try_table` pushes one label
;; but one handler per clause, so disposing "one handler per exited label"
;; leaves every clause after the first armed, and the leak is only observable
;; when something throws AFTER the region was left.

(module
  (tag $a (param i32))
  (tag $b (param i32))
  ;; Same signature as $a — identity, not shape, is what matches.
  (tag $a_twin (param i32))
  (tag $pair (param i32 i64))
  (tag $bare)

  (func $throw_a (param i32) (throw $a (local.get 0)))
  (func $throw_b (param i32) (throw $b (local.get 0)))
  (func $throw_twin (param i32) (throw $a_twin (local.get 0)))

  ;; First matching clause wins, and a later clause for the same tag is dead.
  (func (export "first_clause_wins") (result i32)
    (block $on_second (result i32)
      (block $on_first (result i32)
        (try_table (result i32) (catch $a $on_first) (catch $a $on_second)
          (call $throw_a (i32.const 7))
          (i32.const -1))
        (return))
      (return (i32.add (i32.const 100))))
    (i32.add (i32.const 200)))

  ;; A tag with an identical signature is a different tag.
  ;; $a and $a_twin have the SAME signature. Matching is by identity, so a
  ;; throw of $a_twin must skip the $a clause and take its own: 1005. Were
  ;; tags compared by shape, the $a clause would swallow it and give 105.
  (func (export "twin_tag_not_caught") (result i32)
    (block $on_twin (result i32)
      (block $on_a (result i32)
        (try_table (result i32) (catch $a $on_a) (catch $a_twin $on_twin)
          (call $throw_twin (i32.const 5))
          (i32.const -1))
        (return))
      (return (i32.add (i32.const 100))))
    (i32.add (i32.const 1000)))

  ;; catch_all runs for any tag and receives NO payload.
  (func (export "catch_all_no_payload") (result i32)
    (block $any
      (try_table (catch_all $any)
        (call $throw_b (i32.const 9)))
      (return (i32.const 0)))
    (i32.const 42))

  ;; Clause order decides: catch_all written first shadows a later catch.
  (func (export "catch_all_first_shadows") (result i32)
    (block $any
      (block $specific (result i32)
        (try_table (catch_all $any) (catch $a $specific)
          (call $throw_a (i32.const 3)))
        (return (i32.const 0)))
      (return (i32.add (i32.const 500))))
    (i32.const 1))

  ;; A multi-value tag delivers every payload value, in order.
  (func (export "multi_payload") (result i64)
    (local $i i32)
    (local $l i64)
    (block $caught (result i32 i64)
      (try_table (result i32 i64) (catch $pair $caught)
        (throw $pair (i32.const 4) (i64.const 100)))
      (return (i64.const -1)))
    ;; payload arrives in order, so the i64 is on top
    (local.set $l)
    (local.set $i)
    (i64.add (local.get $l) (i64.extend_i32_s (local.get $i))))

  ;; A tag with no parameters delivers nothing.
  (func (export "bare_tag") (result i32)
    (block $caught
      (try_table (catch $bare $caught)
        (throw $bare))
      (return (i32.const 0)))
    (i32.const 77))

  ;; ── The handler extent ────────────────────────────────────────────────
  ;; Falling off the end of the body ends it: the second throw is NOT caught
  ;; by the region that already completed.
  ;; Returns 55 ONLY if a stale handler caught the throw; otherwise the
  ;; exception leaves this function entirely.
  (func $inner_normal_exit (result i32)
    (block $caught (result i32)
      (try_table (catch $a $caught) (catch $b $caught)
        (nop))
      ;; body completed normally — handlers are gone from here on, so this
      ;; $b must escape rather than reach the clauses above
      (call $throw_b (i32.const 1))
      (return (i32.const 0)))
    (drop)
    (i32.const 55))
  (func (export "escapes_after_normal_exit") (result i32)
    (block $outer (result i32)
      (try_table (result i32) (catch $b $outer)
        (call $inner_normal_exit))
      (return))
    ;; reached with the escaped payload (1) → 901 proves it escaped;
    ;; 55 would mean the completed region caught it
    (i32.add (i32.const 900)))

  ;; Branching OUT of the body ends it too. Two clauses, so a disposal that
  ;; removes one handler per label leaves the `catch_all` armed and wrongly
  ;; swallows the later throw.
  (func $inner_branch_out (result i32)
    (block $caught (result i32)
      (block $left
        (try_table (catch $a $caught) (catch $b $caught)
          (br $left))
        (unreachable))
      ;; reached via $left, outside the try_table's extent — TWO clauses were
      ;; installed, so a disposal that removes one handler per LABEL leaves
      ;; the $b clause armed and wrongly swallows this
      (call $throw_b (i32.const 2))
      (return (i32.const 0)))
    (drop)
    (i32.const 66))
  (func (export "escapes_after_branch_out") (result i32)
    (block $outer (result i32)
      (try_table (result i32) (catch $b $outer)
        (call $inner_branch_out))
      (return))
    (i32.add (i32.const 900)))

  ;; A nested try_table's handlers do not outlive it either.
  (func (export "nested_inner_does_not_leak") (result i32)
    (block $outer_caught (result i32)
      (try_table (result i32) (catch $b $outer_caught)
        (block $inner_done
          (try_table (catch_all $inner_done) (catch_all $inner_done)
            (br $inner_done)))
        ;; inner extent over; $b must reach the OUTER handler
        (call $throw_b (i32.const 8))
        (i32.const -1))
      (return))
    (i32.add (i32.const 0)))

  ;; ── throw_ref preserves tag and payload ──────────────────────────────
  (func (export "throw_ref_keeps_tag") (result i32)
    (local $e exnref)
    (block $outer_a (result i32)
      (block $captured (result exnref)
        (try_table (result exnref) (catch_all_ref $captured)
          (call $throw_a (i32.const 21))
          (unreachable))
        (unreachable))
      (local.set $e)
      ;; re-raise it; the OUTER $a handler must see the ORIGINAL tag, so the
      ;; payload it delivers is the original 21.
      (try_table (catch $a $outer_a)
        (throw_ref (local.get $e)))
      (unreachable))
    (i32.add (i32.const 0)))

  ;; catch_ref delivers the payload AND the exnref.
  (func (export "catch_ref_gives_both") (result i32)
    (block $caught (result i32 exnref)
      (try_table (result i32 exnref) (catch_ref $a $caught)
        (call $throw_a (i32.const 11))
        (unreachable))
      (return (i32.const -1)))
    (drop)
    (i32.add (i32.const 0)))

  ;; An uncaught exception propagates out of the call that raised it.
  (func (export "propagates_through_frames") (result i32)
    (block $caught (result i32)
      (try_table (result i32) (catch $a $caught)
        (call $indirect_thrower)
        (i32.const -1))
      (return))
    (i32.add (i32.const 0)))
  (func $indirect_thrower (call $throw_a (i32.const 13)))
)

;; ── clause order and tag identity ───────────────────────────────────────
(assert_return (invoke "first_clause_wins") (i32.const 107))
(assert_return (invoke "twin_tag_not_caught") (i32.const 1005))
(assert_return (invoke "catch_all_no_payload") (i32.const 42))
(assert_return (invoke "catch_all_first_shadows") (i32.const 1))
(assert_return (invoke "multi_payload") (i64.const 104))
(assert_return (invoke "bare_tag") (i32.const 77))

;; ── handler extent: a region already left must not catch ───────────────
;; Both of these throw $b with no live handler for it, so the exception must
;; escape the function rather than be swallowed.
;; 901/902 = the exception escaped the region, as the spec requires.
;; 55/66 would mean a region already left still caught it.
(assert_return (invoke "escapes_after_normal_exit") (i32.const 901))
(assert_return (invoke "escapes_after_branch_out") (i32.const 902))
(assert_return (invoke "nested_inner_does_not_leak") (i32.const 8))

;; ── throw_ref and catch_ref ────────────────────────────────────────────
(assert_return (invoke "throw_ref_keeps_tag") (i32.const 21))
(assert_return (invoke "catch_ref_gives_both") (i32.const 11))
(assert_return (invoke "propagates_through_frames") (i32.const 13))
