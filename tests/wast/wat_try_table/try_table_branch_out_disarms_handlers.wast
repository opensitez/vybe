;; vybe-test: wast/wat_try_table/try_table_branch_out_disarms_handlers
;; vybe-test-mode: run
;;
;; Leaving a `try_table` by BRANCHING out must disarm the whole protected
;; region, exactly as falling off its `end` does. The spec has one notion of a
;; try_table's extent — the block — and both exits close the same block.
;;
;; This is invisible with ONE catch clause, which is all any existing test
;; uses: a `try_table` installs one handler per CLAUSE but pushes one structural
;; label, so an implementation that disposes of "one handler per exited try
;; label" is right by coincidence at one clause and wrong at two.
;;
;; A stale handler is not a quiet leak. It is still armed, so a later throw that
;; must escape gets caught by a region the program already left, and control
;; resumes at a label belonging to a dead block — so the code after that label
;; runs a SECOND time. Two observables here, because they fail differently:
;;
;;   * a trailing `catch_all` left armed swallows any later throw outright, so
;;     the escape never happens at all;
;;   * a stale TYPED clause only hijacks control — the tail re-runs and the
;;     throw is re-raised, so it does escape in the end. `assert_exception`
;;     cannot see that; a counter can, and must read exactly one.
;;
;; Everything is written FOLDED, including `(drop (block …))`. A plain
;; instruction followed by a folded one — `drop (throw …)` — is mis-parsed here
;; (the folded instruction is absorbed as the plain one's operand); that is a
;; separate front-end defect with its own repro, and this file is about handler
;; disposal, so it stays clear of it.
;;
;; Spec-format so `wasmtime wast` arbitrates.

(module
  (tag $a (param i32))
  (tag $b (param i32))
  (tag $c (param i32))

  (global $n (mut i32) (i32.const 0))
  (func (export "n") (result i32) (global.get $n))

  ;; ── The region still catches while it is live (positive control) ─────
  ;; Without this, "nothing is ever caught" would satisfy the whole file.
  (func (export "live_region_catches_first") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $a $h) (catch $b $h)
        (throw $a (i32.const 11)))))

  (func (export "live_region_catches_second") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $a $h) (catch $b $h)
        (throw $b (i32.const 22)))))

  ;; ── THE discriminating case ──────────────────────────────────────────
  ;; Branch out of a two-clause region, then throw the LAST clause's tag —
  ;; the clause a one-per-label disposal leaves behind (clauses are pushed
  ;; in reverse so the first is tried first). An outer region catches it
  ;; either way, so the answer is the same; what differs is how many times
  ;; the code after the dead branch target ran.
  (func (export "run_once_after_branch_target") (result i32)
    (global.set $n (i32.const 0))
    (block $h (result i32)
      (try_table (result i32) (catch $b $h)
        (drop
          (block $out (result i32)
            (try_table (result i32) (catch $a $out) (catch $b $out)
              (br $out (i32.const 7)))))
        (global.set $n (i32.add (global.get $n) (i32.const 1)))
        (throw $b (i32.const 2))
        (i32.const -1))))

  ;; Three clauses, throwing the middle one — same shape, deeper group.
  (func (export "run_once_three_clauses") (result i32)
    (global.set $n (i32.const 0))
    (block $h (result i32)
      (try_table (result i32) (catch $b $h)
        (drop
          (block $out (result i32)
            (try_table (result i32) (catch $a $out) (catch $b $out) (catch $c $out)
              (br $out (i32.const 7)))))
        (global.set $n (i32.add (global.get $n) (i32.const 1)))
        (throw $b (i32.const 3))
        (i32.const -1))))

  ;; ── `catch_all` as a trailing clause: it swallows ANY later throw ─────
  (func (export "escapes_past_catch_all")
    (block $any
      (drop
        (block $out (result i32)
          (try_table (result i32) (catch $a $out) (catch_all $any)
            (br $out (i32.const 7)))))
      (throw $b (i32.const 5))))

  ;; ── Branching out of the INNER region leaves the OUTER one armed ─────
  (func (export "outer_still_catches") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $b $h) (catch $c $h)
        (drop
          (block $inner (result i32)
            (try_table (result i32) (catch $a $inner) (catch $c $inner)
              (br $inner (i32.const 7)))))
        (throw $b (i32.const 33))
        (i32.const -1))))

  ;; ── A throw in a CALLEE after the caller branched out ────────────────
  (func $raise (throw $b (i32.const 44)))
  (func (export "escapes_from_callee")
    (drop
      (block $out (result i32)
        (try_table (result i32) (catch $a $out) (catch $b $out)
          (br $out (i32.const 7)))))
    (call $raise))
)

(assert_return (invoke "live_region_catches_first") (i32.const 11))
(assert_return (invoke "live_region_catches_second") (i32.const 22))
(assert_return (invoke "outer_still_catches") (i32.const 33))

(assert_return (invoke "run_once_after_branch_target") (i32.const 2))
(assert_return (invoke "n") (i32.const 1))

(assert_return (invoke "run_once_three_clauses") (i32.const 3))
(assert_return (invoke "n") (i32.const 1))

(assert_exception (invoke "escapes_past_catch_all"))
(assert_exception (invoke "escapes_from_callee"))
