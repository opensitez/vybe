;; vybe-test: wast/wat_br_on_cast/branch_carries_the_values_below_the_reference
;; vybe-test-mode: run
;;
;; `br_on_cast $l rt1 rt2` is typed `t* rt_1 -> t* (rt_1 \ rt_2)`. That leading
;; `t*` is real: the branch delivers EVERY result of the target block, not just
;; the reference on top. A block declared `(result i32 i64 anyref)` branched
;; out of by `br_on_cast` must carry the i32 and the i64 sitting below the
;; reference.
;;
;; Assigning only the topmost result left those unwritten, so the block
;; delivered whatever the temps happened to hold — a WRONG VALUE, never a
;; validation error, which is why nothing upstream caught it.
;;
;; ⛔ THE PROPOSAL'S OWN `br_on_cast` SUITE CANNOT SEE THIS. Every block in it
;; is single-result, so `t*` is empty in every case and the drop is invisible.
;; It surfaced from the Custom Descriptors suite, whose "Sent values" section
;; branches out of a multi-result block — and `br_on_cast_desc_eq` was fixed
;; there while the plain form, which shares its shape and none of its code, was
;; not. Each sibling is spelled out below for that reason: they are separate
;; emitters, and being right is not transitive between them.
;;
;; Results are taken into LOCALS rather than compared on the stack: `i32.eq`
;; consumes the two topmost values, so a stack-only check reads the operands it
;; is trying to distinguish and can pass while the carry is wrong.

(module
  (type $a (struct (field i32)))
  (type $b (struct (field i64)))

  ;; ── the reference casts SUCCESSFULLY: the branch is taken ──────────
  (func (export "taken") (result i32)
    (local $x i32) (local $y i64)
    (block $l (result i32 i64 anyref)
      (i32.const 11)
      (i64.const 22)
      (struct.new $a (i32.const 7))
      (br_on_cast $l anyref (ref $a))
    )
    (drop)             ;; the reference
    (local.set $y)     ;; the i64 below it
    (local.set $x)     ;; …and the i32 below that
    (i32.and
      (i64.eq (local.get $y) (i64.const 22))
      (i32.eq (local.get $x) (i32.const 11))))

  ;; ── the cast FAILS: the branch is NOT taken, the block falls out ────
  ;; The same three values reach the end of the block by the ordinary path.
  ;; This is the control proving a carried `t*` was not consumed off the
  ;; stack — the fall-through still needs it.
  (func (export "not_taken") (result i32)
    (local $x i32) (local $y i64)
    (block $l (result i32 i64 anyref)
      (i32.const 11)
      (i64.const 22)
      (struct.new $b (i64.const 9))
      (br_on_cast $l anyref (ref $a))
    )
    (drop)
    (local.set $y)
    (local.set $x)
    (i32.and
      (i64.eq (local.get $y) (i64.const 22))
      (i32.eq (local.get $x) (i32.const 11))))

  ;; ── ORDER, not just presence ───────────────────────────────────────
  ;; Two i32s below the reference: a carry that wrote the temps in the wrong
  ;; direction still writes BOTH, and only a positional check tells them apart.
  (func (export "order") (result i32)
    (local $x i32) (local $y i32)
    (block $l (result i32 i32 anyref)
      (i32.const 100)
      (i32.const 200)
      (struct.new $a (i32.const 7))
      (br_on_cast $l anyref (ref $a))
    )
    (drop)
    (local.set $y)     ;; topmost of the two ⇒ 200
    (local.set $x)     ;; the one below it   ⇒ 100
    (i32.and
      (i32.eq (local.get $y) (i32.const 200))
      (i32.eq (local.get $x) (i32.const 100))))

  ;; ── a DEEPER `t*` ──────────────────────────────────────────────────
  ;; Four values below the reference, all distinct. One temp landing by luck
  ;; proves nothing; this fails unless the carry is positional across the whole
  ;; run.
  (func (export "deep") (result i32)
    (local $p i32) (local $q i32) (local $r i32) (local $s i32)
    (block $l (result i32 i32 i32 i32 anyref)
      (i32.const 1) (i32.const 2) (i32.const 3) (i32.const 4)
      (struct.new $a (i32.const 7))
      (br_on_cast $l anyref (ref $a))
    )
    (drop)
    (local.set $s) (local.set $r) (local.set $q) (local.set $p)
    (i32.add
      (i32.add (i32.mul (local.get $p) (i32.const 1000))
               (i32.mul (local.get $q) (i32.const 100)))
      (i32.add (i32.mul (local.get $r) (i32.const 10))
               (local.get $s))))

  ;; ── the PLAIN (flat) spelling ──────────────────────────────────────
  ;; ⛔ FOLDED ≠ PLAIN. `emit_br_on_cast_stmt` is reached from two parse sites,
  ;; and the carry reads whatever is live on the walker's operand stack at that
  ;; moment — which the folded and flat walkers arrive at differently (the
  ;; folded one drains and rebinds across a block boundary). A fix verified in
  ;; one spelling says nothing about the other; that is exactly how the memidx
  ;; peel was half-fixed.
  (func (export "taken_plain") (result i32)
    (local $x i32) (local $y i64)
    block $l (result i32 i64 anyref)
      i32.const 11
      i64.const 22
      (struct.new $a (i32.const 7))
      br_on_cast $l anyref (ref $a)
    end
    drop
    local.set $y
    local.set $x
    (i32.and
      (i64.eq (local.get $y) (i64.const 22))
      (i32.eq (local.get $x) (i32.const 11))))

  ;; ── the siblings, on the same multi-result shape ───────────────────
  (func (export "fail_taken") (result i32)
    (local $x i32) (local $y i64)
    (block $l (result i32 i64 anyref)
      (i32.const 33)
      (i64.const 44)
      (struct.new $b (i64.const 9))
      (br_on_cast_fail $l anyref (ref $a))   ;; not an $a ⇒ branch
    )
    (drop)
    (local.set $y)
    (local.set $x)
    (i32.and
      (i64.eq (local.get $y) (i64.const 44))
      (i32.eq (local.get $x) (i32.const 33))))

  (func (export "brif") (result i32)
    (local $x i32) (local $y i64)
    (block $l (result i32 i64 anyref)
      (i32.const 55)
      (i64.const 66)
      (struct.new $a (i32.const 7))
      (br_if $l (i32.const 1))
    )
    (drop)
    (local.set $y)
    (local.set $x)
    (i32.and
      (i64.eq (local.get $y) (i64.const 66))
      (i32.eq (local.get $x) (i32.const 55))))
)

(assert_return (invoke "taken") (i32.const 1))
(assert_return (invoke "not_taken") (i32.const 1))
(assert_return (invoke "order") (i32.const 1))
(assert_return (invoke "deep") (i32.const 1234))
(assert_return (invoke "taken_plain") (i32.const 1))
(assert_return (invoke "fail_taken") (i32.const 1))
(assert_return (invoke "brif") (i32.const 1))
