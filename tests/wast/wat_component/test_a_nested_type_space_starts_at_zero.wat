;; vybe-test: wast/wat_component/test_a_nested_type_space_starts_at_zero
;; hand-written against proposals/component-model/design/mvp/Explainer.md §8
;;   — every component has its OWN type index space, numbered from 0.
;;
;; ▶▶ TWO NUMBERINGS HAVE TO COEXIST, AND THEY DISAGREED SILENTLY.
;;
;;   the SOURCE numbers types per component, from 0     (the spec)
;;   the VM has ONE `canon_types` table for the program (one flat vector)
;;
;; So a nested component's `(type $t1)` is index 0 to the source and index 1 in
;; the payload vector when the enclosing component already declared one. The
;; walker kept a per-component COUNTER and a shared VECTOR and let them drift:
;; the counter reset to 0 on entry while the vector kept growing.
;;
;; ⛔ THE DEBUG BUILD CAUGHT IT AS `assertion left == right: 2 vs 1`. A RELEASE
;; BUILD WOULD NOT HAVE. `debug_assert_eq!` compiles away, and then
;; `canon lift (type $t1)` reads `comp_types[0]` — the OUTER component's type —
;; and lifts with a signature the source never wrote. Both are small integers
;; and both are in range, so nothing downstream can tell.
;;
;; The fix is a BASE OFFSET per component, applied in exactly one resolver:
;; a `$name` is stored already-global, a bare integer is rebased. Applying it
;; at each call site is the shape that has gone wrong in this tree before —
;; `case_sensitive` needed `!self.case_sensitive &&` at 33 sites and 23 forgot.
;;
;; THE DISCRIMINATOR IS THAT THE OUTER TYPE EXISTS AND IS NEVER USED. Remove
;; `(type $outer …)` and both numberings agree by accident, which is exactly
;; the configuration every earlier nested-component test happened to have —
;; which is why none of them caught this.
;;
;; The inner returns 1, so a lift that read the outer's `(result u32)` would
;; still typecheck and still return something; only a signature mismatch or a
;; trap would show. The assert is the cheap half; the header is the record of
;; what it is really guarding.

(component
  ;; Declared, never referenced. Its only job is to make the two numberings
  ;; differ by one.
  (type $outer (func (result u32)))

  (component $inner
    (type $t1 (func (result u32)))
    (core module $m
      (func (export "one") (result i32) (i32.const 1)))
    (core instance $mi (instantiate $m))
    (alias core export $mi "one" (core func $d))
    (canon lift (core func $d) (func $lifted (type $t1)))
    (export "one" (func $lifted))
  )

  (instance $i (instantiate $inner))
  (alias export $i "one" (func $r))
  (canon lower (func $r) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (result i32)))
    (func (export "get") (result i32) (call $l)))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 1))
