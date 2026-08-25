;; vybe-test: wast/wat_component/test_two_core_instances_are_independent
;; hand-written against proposals/component-model/design/mvp/Explainer.md §3
;;   `core:instanceexpr ::= (instantiate <core:moduleidx> <core:instantiatearg>*)`
;;
;; Instantiating one module TWICE is the point of `instantiate`, and the two
;; instances must not share state. The discriminator is a mutable global that
;; the module's own `start` increments:
;;
;;   * two independent instances → each has its own `$g`, its own `start` ran
;;     once, and `get` reads **1**;
;;   * both instantiations collapsed onto one → one `$g`, `start` ran twice, and
;;     `get` reads **2**.
;;
;; This is worth pinning because the collapse would be SILENT. The walker mints
;; a class per `walk_module` call but publishes into registries keyed by module
;; name, and an assert that only checked `get` returned *something* would pass
;; either way.

(component
  (core module $m
    (global $g (mut i32) (i32.const 0))
    (func $bump
      (global.set $g (i32.add (global.get $g) (i32.const 1))))
    (start $bump)
    (func (export "get") (result i32)
      (global.get $g)))
  (core instance (instantiate $m))
  (core instance (instantiate $m))
)

(assert_return (invoke "get") (i32.const 1))
