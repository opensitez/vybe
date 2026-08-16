;; vybe-test: wast/wat_try_table/throw_ref_rethrows_the_same_exception
;; vybe-test-mode: run
;;
;; `throw_ref` re-raises an exception CAPTURED as an `exnref`. It is not "throw
;; something": the re-raised exception must be the same entity — same tag, same
;; payload — so a clause further out that names the original tag still matches,
;; and one that names a different tag still does not.
;;
;; `throw_ref` occurred once in the whole run corpus. Everything below is what
;; that one occurrence could not distinguish:
;;
;;   * the rethrown exception keeps its TAG (an implementation that re-raises
;;     under the generic exception tag is caught by a `catch_all` and looks
;;     fine, but a typed clause further out never fires);
;;   * it keeps its PAYLOAD across the capture/rethrow round trip;
;;   * `catch_ref` delivers the payload values FIRST and the exnref on top —
;;     the order matters and is invisible with a zero-arity tag;
;;   * a rethrow crosses a frame boundary like any other throw;
;;   * `throw_ref` on a NULL exnref traps.
;;
;; Spec-format so `wasmtime wast` arbitrates.

(module
  (tag $a (param i32))
  (tag $b (param i32))

  ;; ── The rethrown exception still matches its own tag ─────────────────
  (func (export "rethrow_keeps_tag") (result i32)
    (block $outer (result i32)
      (try_table (result i32) (catch $a $outer)
        (block $caught (result exnref)
          (try_table (catch_all_ref $caught)
            (throw $a (i32.const 42)))
          (return (i32.const -1)))
        (throw_ref))))

  ;; ── …and does NOT match a different one ──────────────────────────────
  ;; A rethrow that lost its tag would be caught here.
  (func (export "rethrow_wrong_tag_escapes")
    (block $caught (result exnref)
      (try_table (catch_all_ref $caught)
        (throw $b (i32.const 1)))
      (return))
    (throw_ref))

  (func (export "rethrow_wrong_tag_not_caught_by_a") (result i32)
    (block $outer (result i32)
      (try_table (result i32) (catch $a $outer)
        (block $caught (result exnref)
          (try_table (catch_all_ref $caught)
            (throw $b (i32.const 1)))
          (return (i32.const -1)))
        (throw_ref))))

  ;; ── `catch_ref` delivers payload THEN exnref ─────────────────────────
  (func (export "catch_ref_payload_order") (result i32) (local $e exnref)
    (block $caught (result i32 exnref)
      (try_table (result i32 exnref) (catch_ref $a $caught)
        (throw $a (i32.const 7))))
    (local.set $e))

  ;; The payload survives capture and rethrow.
  (func (export "rethrow_keeps_payload") (result i32) (local $e exnref)
    (block $outer (result i32)
      (try_table (result i32) (catch $a $outer)
        (block $caught (result i32 exnref)
          (try_table (result i32 exnref) (catch_ref $a $caught)
            (throw $a (i32.const 1234))))
        (local.set $e)
        (drop)
        (throw_ref (local.get $e)))))

  ;; ── A rethrow crosses a frame boundary ───────────────────────────────
  (func $rethrower (param $e exnref) (throw_ref (local.get $e)))
  (func (export "rethrow_from_callee") (result i32)
    (block $outer (result i32)
      (try_table (result i32) (catch $a $outer)
        (block $caught (result exnref)
          (try_table (catch_all_ref $caught)
            (throw $a (i32.const 99)))
          (return (i32.const -1)))
        (call $rethrower)
        (i32.const -2))))

  ;; ── A null exnref traps ──────────────────────────────────────────────
  (func (export "throw_ref_null_traps") (throw_ref (ref.null exn)))
)

(assert_return (invoke "rethrow_keeps_tag") (i32.const 42))
(assert_return (invoke "rethrow_keeps_payload") (i32.const 1234))
(assert_return (invoke "catch_ref_payload_order") (i32.const 7))
(assert_return (invoke "rethrow_from_callee") (i32.const 99))

(assert_exception (invoke "rethrow_wrong_tag_escapes"))
(assert_exception (invoke "rethrow_wrong_tag_not_caught_by_a"))

(assert_trap (invoke "throw_ref_null_traps") "null exception reference")
