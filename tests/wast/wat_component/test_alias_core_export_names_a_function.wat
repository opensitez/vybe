;; vybe-test: wast/wat_component/test_alias_core_export_names_a_function
;; hand-written against proposals/component-model/design/mvp/Explainer.md §5
;;   `alias ::= (alias core export <core:instanceidx> <core:name>
;;                   (core <core:sort> <id>?))`
;;
;; `(alias core export …)` is the only way a component reaches INSIDE an
;; instantiated core module, and it is the second producer of the core
;; function index space — the first being a canon row's binder.
;;
;; Both producers now feed ONE space, which is what the spec requires: a
;; positional `(func 0)` and a named `(func $a)` must be able to name the same
;; entity. Before this the space did not exist at all — `core_func_index` was
;; read in six places and written in none, so every `$id` in it reported "not
;; bound in the core func index space" no matter what the source declared.
;; That empty space, not a missing walk, is why `canon lift` could not bind a
;; callee.
;;
;; The alias binds `$a`, and the `with` clause supplies it under the name the
;; importing module chose ("f"), which is deliberately NOT the name the
;; exporting module published ("answer") — so a green result cannot come from
;; the names happening to match.

(component
  (core module $lib
    (func (export "answer") (result i32)
      (i32.const 7)))
  (core instance $li (instantiate $lib))

  (alias core export $li "answer" (core func $a))

  (core module $main
    (import "l" "f" (func $f (result i32)))
    (func (export "get") (result i32)
      (call $f)))
  (core instance (instantiate $main
    (with "l" (instance (export "f" (func $a))))))
)

(assert_return (invoke "get") (i32.const 7))
