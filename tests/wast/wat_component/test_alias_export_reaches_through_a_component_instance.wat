;; vybe-test: wast/wat_component/test_alias_export_reaches_through_a_component_instance
;; hand-written against proposals/component-model/design/mvp/Explainer.md
;;   §Instance definitions and §Alias definitions:
;;     (instance <id>? (export <name> <externidx>)*)
;;     (alias export <instanceidx> <name> (<sort> <id>?))
;;
;; ▶▶ THE COMPONENT INSTANCE INDEX SPACE, and the SECOND producer of the
;; component FUNCTION index space.
;;
;;   lift ×2  →  (instance (export "op" …))  →  (alias export …)  →  lower
;;
;; Until now `canon lift` was the only thing that could put an entry in the
;; component function space, so `canon lower` could only ever name a function
;; lifted a few lines above it. An alias reaches a function through an
;; instance's export table, which is how a component addresses anything it did
;; not define itself.
;;
;; ⛔ THE DISCRIMINATOR IS THAT THERE ARE **TWO** LIFTS. With one lifted
;; function an alias that resolved to the wrong index — 0 instead of 1, or the
;; aliased entry instead of the alias — would land back on the same function
;; and the test would pass anyway. Here `$times2` is component func 0 and
;; `$times3` is 1; the instance exports the SECOND. A resolution that answers 0
;; returns 28 rather than 42, and one that reads an empty export table refuses
;; by name.
;;
;; `instanceexpr` is a rule of its own rather than a silent one, so the
;; `(export …)` pairs are GRANDCHILDREN of `(instance …)`. Reading them as
;; direct children finds none, publishes an empty instance, and makes this file
;; fail with `exports no "op"` — a walk bug that reads as a source error.

(component
  (core module $m
    (func (export "times2") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2)))
    (func (export "times3") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 3))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "times2" (core func $c2))
  (alias core export $mi "times3" (core func $c3))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $c2) (func $times2 (type $ft)))
  (canon lift (core func $c3) (func $times3 (type $ft)))

  (instance $i (export "op" (func $times3)))
  (alias export $i "op" (func $chosen))

  (canon lower (func $chosen) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 14))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
