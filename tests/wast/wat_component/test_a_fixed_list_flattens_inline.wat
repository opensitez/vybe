;; vybe-test: wast/wat_component/test_a_fixed_list_flattens_inline
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Flattening (:3205)  flatten_list — `flatten_type(t) * N` when N is present
;;   §Storing   (:2997)  store_list   — `assert(N == len(v))`, elements INLINE
;;   §Alignment (:2248)  alignment_list — the ELEMENT's alignment, not a pointer's
;;
;; ▶▶ 🔧 A FIXED LIST IS A DIFFERENT SHAPE, NOT A `list` WITH A CONSTRAINT.
;; An unfixed `list` is a `(ptr, length)` pair — TWO core values, elements
;; elsewhere in memory, `realloc` called to put them there. A fixed one is
;; INLINE: `list<u32, 3>` occupies THREE core `i32`s and no pointer at all.
;;
;; THE ARITY IS THE PROOF. The core callee declares `(param i32 i32 i32)` and
;; the lowered core function does too. Had `list<u32,3>` been modelled as an
;; unfixed `list`, it would flatten to two pointer-sized values and this would
;; fail on ARITY before producing any number — and had it been modelled as a
;; single value, on arity again. Only `flatten_type(u32) * 3` gives three.
;;
;; The callee answers `a*100 + b*10 + c` = 123, so the three elements must
;; arrive in ORDER as well as in the right count: 321 would mean reversed, and
;; any repeated element gives 111 / 222 / 333.
;;
;; ⛔ A `list<T, 0>` is REFUSED at compile time, not accepted as an empty list:
;; it would flatten to NO core values, so the parameter would vanish from the
;; signature and every argument after it would shift by one. That is the same
;; class as the 33rd flag having no bit — an absence nothing downstream can
;; detect.

(component
  (core module $m
    (func (export "digits") (param i32 i32 i32) (result i32)
      (i32.add
        (i32.add (i32.mul (local.get 0) (i32.const 100))
                 (i32.mul (local.get 1) (i32.const 10)))
        (local.get 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "digits" (core func $d))

  (type $ft (func (param "xs" (list u32 3)) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))
  (canon lower (func $lifted) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32 i32 i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 1) (i32.const 2) (i32.const 3))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 123))
