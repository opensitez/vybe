;; vybe-test: wast/wat_component/test_an_alias_defines_a_new_component_func_index
;; hand-written against proposals/component-model/design/mvp/Explainer.md
;;   §Alias definitions — "each alias definition … adds a new index to the
;;   index space of the aliased sort".
;;
;; ▶▶ AN ALIAS APPENDS. It does not give the aliased entity a second name.
;;
;; The component function index space here is:
;;
;;   funcidx 0 → the `canon lift`
;;   funcidx 1 → the ALIAS of it, reached through instance $i
;;
;; and `canon lower` names funcidx **1 POSITIONALLY**. That spelling is the
;; whole point: a `$id`-only test cannot tell an appending alias from a
;; renaming one, because both bind the id to something callable. If the alias
;; only rebound a name, funcidx 1 would not exist and this refuses with
;; `component func 1 is not defined (have 1)`.
;;
;; It matters beyond bookkeeping — every later positional reference in the
;; component counts on the space having advanced. An alias that silently did
;; not advance it would shift every subsequent index by one, which reads as a
;; correct program right up until two functions have the same signature.
;;
;; 21 × 2 = 42, so a lower that reached nothing would return 21 and one that
;; dropped its argument would return 0.

(component
  (core module $m
    (func (export "double") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "double" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))

  (instance $i (export "run" (func $lifted)))
  (alias export $i "run" (func))

  (canon lower (func 1) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 21))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
