;; vybe-test: wast/wat_try_table/try_table_label_depths_and_clause_kinds
;; origin: the labelidx-as-byte-offset defect — nothing exercised a depth > 0
;; vybe-test-mode: run
;;
;; A catch clause's handler target is a spec `labelidx`: a RELATIVE BLOCK DEPTH,
;; resolved like `br`'s operand. It is not a byte offset, and the difference is
;; only visible when the depth is greater than zero — every existing test in
;; `wat_try_table/` catches into the block immediately enclosing the
;; `try_table`, which a wrong-but-consistent encoding also satisfies.
;;
;; So the cases that matter here are:
;;
;;   * a clause branching PAST one or more intervening blocks (depth 1, 2, 3);
;;   * the payload arriving as the TARGET block's result — a branch keeps
;;     exactly the target's result arity, so a target typed too narrowly
;;     silently drops values;
;;   * all four clause kinds, which differ precisely in what they deliver:
;;     `catch` the payload, `catch_ref` payload + exnref, `catch_all` nothing,
;;     `catch_all_ref` the exnref only;
;;   * clause ORDER — first match wins, so a `catch_all` written first must
;;     shadow a later typed clause, and written last must not;
;;   * tag IDENTITY — two tags with the same signature never match each other.
;;
;; Spec-format so `wasmtime wast` arbitrates every one of these.

(module
  (tag $a (param i32))
  (tag $b (param i32))
  (tag $two (param i32 i32))

  ;; ── depth 0: the shape every existing test already covers ────────────
  (func (export "depth0") (result i32)
    (block $h (result i32)
      (try_table (catch $a $h)
        (throw $a (i32.const 11)))
      (i32.const -1)))

  ;; ── depth 1, 2, 3: the clause branches PAST intervening blocks ───────
  ;; A byte offset reaches any of these equally well, which is why they were
  ;; the blind spot; a depth must count the blocks correctly.
  (func (export "depth1") (result i32)
    (block $out (result i32)
      (block $skip
        (try_table (catch $a $out)
          (throw $a (i32.const 21))))
      (i32.const -1)))

  (func (export "depth2") (result i32)
    (block $out (result i32)
      (block $mid
        (block $inner
          (try_table (catch $a $out)
            (throw $a (i32.const 32)))))
      (i32.const -1)))

  (func (export "depth3") (result i32)
    (block $out (result i32)
      (block $b3
        (block $b2
          (block $b1
            (try_table (catch $a $out)
              (throw $a (i32.const 43))))))
      (i32.const -1)))

  ;; The INNERMOST target still works when outer blocks exist — the depth must
  ;; be 0 here, not "the outermost handler".
  (func (export "innermost_wins") (result i32)
    (block $out (result i32)
      (block $near (result i32)
        (try_table (catch $a $near)
          (throw $a (i32.const 51)))
        (i32.const -1))))

  ;; ── a LOOP label is a legal target too, and it is not a block ────────
  ;; Branching to a loop jumps to its START, so this would spin forever if the
  ;; clause targeted the loop; targeting the block past it is the terminating
  ;; case. Included so the depth walk is exercised across a mixed label stack.
  (func (export "past_a_loop") (result i32)
    (block $out (result i32)
      (loop $l
        (try_table (catch $a $out)
          (throw $a (i32.const 61))))
      (i32.const -1)))

  ;; ── the four clause kinds ────────────────────────────────────────────
  ;; `catch` delivers the tag's payload.
  (func (export "kind_catch") (result i32)
    (block $h (result i32)
      (try_table (catch $a $h)
        (throw $a (i32.const 71)))
      (i32.const -1)))

  ;; Multi-value payload: BOTH values arrive, in order. A target typed with one
  ;; result would keep only the last.
  (func (export "kind_catch_multi") (result i32)
    (block $h (result i32 i32)
      (try_table (catch $two $h)
        (throw $two (i32.const 7) (i32.const 9)))
      (i32.const -1)
      (i32.const -1))
    (i32.sub))

  ;; `catch_all` matches any tag and delivers NOTHING — the target is void.
  (func (export "kind_catch_all") (result i32)
    (block $h
      (try_table (catch_all $h)
        (throw $b (i32.const 81))))
    (i32.const 81))

  ;; `catch_ref` delivers payload + exnref; `throw_ref` re-raises it.
  (func (export "kind_catch_ref") (result i32)
    (block $outer (result i32)
      (block $h (result i32 exnref)
        (try_table (catch_ref $a $h)
          (throw $a (i32.const 91)))
        (i32.const -1)
        (unreachable))
      (drop)          ;; drop the exnref, keep the payload
      (br $outer)))

  ;; `catch_all_ref` delivers only the exnref.
  (func (export "kind_catch_all_ref") (result i32)
    (block $h (result exnref)
      (try_table (catch_all_ref $h)
        (throw $a (i32.const 101)))
      (unreachable))
    (drop)
    (i32.const 101))

  ;; ── clause ORDER: first match wins ───────────────────────────────────
  ;; catch_all FIRST shadows the typed clause after it...
  (func (export "order_all_first") (result i32)
    (block $typed (result i32)
      (block $all
        (try_table (catch_all $all) (catch $a $typed)
          (throw $a (i32.const 1))))
      (i32.const 111))) ;; catch_all won

  ;; ...and the same clauses in the other order let the typed one match.
  ;; `$all` is VOID because `catch_all` delivers no values — its target may not
  ;; have result types, so the two clauses cannot share one target block.
  (func (export "order_typed_first") (result i32)
    (block $all
      (block $typed (result i32)
        (try_table (catch $a $typed) (catch_all $all)
          (throw $a (i32.const 121)))
        (unreachable))
      (return))
    (i32.const -1))

  ;; ── tag IDENTITY, never signature ────────────────────────────────────
  ;; $a and $b have identical signatures; a clause on $a must not catch $b.
  (func (export "identity_no_match") (result i32)
    (block $outer
      (block $wrong (result i32)
        (try_table (catch_all $outer)
          (try_table (catch $a $wrong)
            (throw $b (i32.const 131))))
        (unreachable))
      (drop))
    (i32.const 141))

  ;; ── no throw: the body completes and the handler is skipped ──────────
  (func (export "no_throw") (result i32)
    (block $h (result i32)
      (try_table (catch $a $h)
        (nop))
      (i32.const 151)))

  ;; ── the target may be written as a NUMERIC DEPTH ─────────────────────
  ;; Same quantity a `br` operand names, and the only form the binary encoding
  ;; has. Depth 0 is the innermost enclosing block; depth 1 skips it — which is
  ;; observable only because the skipped block still has work after it.
  (func (export "numeric_depth_0") (result i32)
    (block $outer (result i32)
      (block $inner (result i32)
        (try_table (catch $a 0)
          (throw $a (i32.const 161)))
        (i32.const -1))
      (i32.const 1)
      (i32.add)))

  (func (export "numeric_depth_1") (result i32)
    (block $outer (result i32)
      (block $inner (result i32)
        (try_table (catch $a 1)
          (throw $a (i32.const 171)))
        (i32.const -1))
      (i32.const 1)
      (i32.add)))

  ;; Depth equal to the number of enclosing labels is the FUNCTION's own
  ;; implicit label — the one `return` targets — and the payload travels as the
  ;; function's result.
  (func (export "numeric_depth_function") (result i32)
    (try_table (catch $a 0)
      (throw $a (i32.const 181)))
    (i32.const -1))

  ;; A tag may be named by index as well: `0` is `$a`, the first declaration.
  (func (export "numeric_tagidx") (result i32)
    (block $h (result i32)
      (try_table (catch 0 $h)
        (throw 0 (i32.const 191)))
      (i32.const -1)))
)

;; ── depth is what is being checked ──────────────────────────────────────
(assert_return (invoke "depth0") (i32.const 11))
(assert_return (invoke "depth1") (i32.const 21))
(assert_return (invoke "depth2") (i32.const 32))
(assert_return (invoke "depth3") (i32.const 43))
(assert_return (invoke "innermost_wins") (i32.const 51))
(assert_return (invoke "past_a_loop") (i32.const 61))

;; ── the four clause kinds deliver different things ──────────────────────
(assert_return (invoke "kind_catch") (i32.const 71))
(assert_return (invoke "kind_catch_multi") (i32.const -2))
(assert_return (invoke "kind_catch_all") (i32.const 81))
(assert_return (invoke "kind_catch_ref") (i32.const 91))
(assert_return (invoke "kind_catch_all_ref") (i32.const 101))

;; ── ordering and identity ───────────────────────────────────────────────
(assert_return (invoke "order_all_first") (i32.const 111))
(assert_return (invoke "order_typed_first") (i32.const 121))
(assert_return (invoke "identity_no_match") (i32.const 141))
(assert_return (invoke "no_throw") (i32.const 151))

;; ── numeric label depths and numeric tag indices ────────────────────────
(assert_return (invoke "numeric_depth_0") (i32.const 162))
(assert_return (invoke "numeric_depth_1") (i32.const 171))
(assert_return (invoke "numeric_depth_function") (i32.const 181))
(assert_return (invoke "numeric_tagidx") (i32.const 191))
