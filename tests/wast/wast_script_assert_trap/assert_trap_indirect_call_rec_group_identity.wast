;; vybe-test: wast/wast_script_assert_trap/assert_trap_indirect_call_rec_group_identity
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

;; Iso-recursive identity: a type declared inside a `(rec …)` is identified by
;; its WHOLE GROUP plus its POSITION in it, never by its own shape. Every `$f1`
;; and `$f2` below is spelled `(func)`, so the signatures are equal in all three
;; modules and only canonicalisation can tell them apart.

;; Same group shape, same position — the SAME type. Must NOT trap.
(module
  (rec (type $f1 (func)) (type (struct)))
  (rec (type $f2 (func)) (type (struct)))
  (table funcref (elem $f1))
  (func $f1 (type $f1))
  (func (export "run") (call_indirect (type $f2) (i32.const 0)))
)
(assert_return (invoke "run"))

;; Same members, different ORDER — different types. Must trap.
(module
  (rec (type $f1 (func)) (type (struct)))
  (rec (type (struct)) (type $f2 (func)))
  (table funcref (elem $f1))
  (func $f1 (type $f1))
  (func (export "run") (call_indirect (type $f2) (i32.const 0)))
)
(assert_trap (invoke "run") "indirect call type mismatch")

;; Different group SIZE — different types. Must trap.
(module
  (rec (type $f1 (func)) (type (struct)))
  (rec (type $f2 (func)))
  (table funcref (elem $f1))
  (func $f1 (type $f1))
  (func (export "run") (call_indirect (type $f2) (i32.const 0)))
)
(assert_trap (invoke "run") "indirect call type mismatch")

;; A type outside any `(rec …)` is its own singleton group, so two identical
;; standalone declarations still merge — the group key must not split them.
(module
  (type $s1 (func))
  (type $s2 (func))
  (table funcref (elem $g))
  (func $g (type $s1))
  (func (export "run") (call_indirect (type $s2) (i32.const 0)))
)
(assert_return (invoke "run"))

;; ⛔ AND THE OVER-FIRE. Parameter NAMES are not part of a type: these two are
;; the same type spelled differently, and the call must succeed. Canonicalisation
;; works off the composite's source text, which does not see that — so the trap
;; can only be decided by the rec-group SHAPE (size and position), which comes
;; from the tree and is exact.
(module
  (type $n1 (func (param f32 f32)))
  (type $n2 (func (param $x f32) (param $y f32)))
  (func $h (type $n1))
  (table funcref (elem $h))
  (func (export "run") (call_indirect (type $n2) (f32.const 1) (f32.const 2) (i32.const 0)))
)
(assert_return (invoke "run"))

;; Same, one level of indirection: `$r1` and `$r2` are equivalent, so `$u1` and
;; `$u2` are too, though their texts differ by the referenced name.
(module
  (type $r1 (func (param i32)))
  (type $r2 (func (param i32)))
  (type $u1 (func (param (ref $r1))))
  (type $u2 (func (param (ref $r2))))
  (func $k (type $u1))
  (func $a (type $r1))
  (table funcref (elem $k))
  (func (export "run") (call_indirect (type $u2) (ref.func $a) (i32.const 0)))
)
(assert_return (invoke "run"))
