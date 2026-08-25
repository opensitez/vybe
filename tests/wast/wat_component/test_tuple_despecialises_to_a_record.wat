;; vybe-test: wast/wat_component/test_tuple_despecialises_to_a_record
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Despecialization (:2175):
;;     case TupleType(ts) : return RecordType([ FieldType(str(i), t) … ])
;;
;; ▶▶ `tuple` IS NOT A MISSING `ValType`. It is a SPECIALIZED type, and every
;; layout, lift, lower and flatten rule in the Canonical ABI matches on
;; `despecialize(t)` rather than on `t` — so the spec's own ABI for a tuple IS
;; the record expansion. Expanding it in `lower_valspec` is the definition, not
;; a stand-in for a variant that ought to exist.
;;
;; ⛔ `string` and `flags` are DELIBERATELY absent from `despecialize()`, and
;; the paragraph under it says why: they have representations distinct from
;; their expansions. `flags` bit-packs into ONE integer of 1, 2 or 4 bytes; a
;; record of `bool` would take a byte per flag. So `flags` still refuses, and
;; that asymmetry is the spec's, not an omission here.
;;
;; What this file proves, which a compile-only check could not: `(tuple u32
;; u32)` flattens to TWO i32s. The core callee takes two parameters and the
;; lowered core function does too, so a tuple that flattened to one value — or
;; to the `(ptr, length)` pair a `list` takes — would fail on arity rather than
;; on the answer.
;;
;; 20 + 22 = 42, and the operands differ, so a flattening that passed the same
;; value twice would return 40 or 44.

(component
  (core module $m
    (func (export "add") (param i32 i32) (result i32)
      (i32.add (local.get 0) (local.get 1))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "add" (core func $a))

  (type $ft (func (param "p" (tuple u32 u32)) (result u32)))
  (canon lift (core func $a) (func $summed (type $ft)))
  (canon lower (func $summed) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32 i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 20) (i32.const 22))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
