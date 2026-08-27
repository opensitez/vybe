;; vybe-test: wast/wat_module/export_name_and_id_are_different_namespaces
;; vybe-test-mode: run
;;
;; An EXPORT NAME and a `$id` live in different namespaces, and a module may
;; use the same spelling for both.
;;
;; This front end lowers a module to a class whose METHODS are its functions,
;; and it named an unnamed exported function after its first export. That is
;; fine until some other function is declared with that spelling as its `$id`:
;; then both want the same method, one silently replaces the other, and a
;; `call $get` inside the exported wrapper resolves to THE WRAPPER.
;;
;; The failure is an infinite recursion reported as a stack overflow, with
;; nothing in it naming a function, an export, or a collision. `gc/array.wast`
;; is built entirely out of this pair — a private `$f` taking the real
;; arguments, and an exported wrapper of the same name that supplies them —
;; so the whole file died on the first invoke.
;;
;; The fix keeps the export reachable by its name while declaring the wrapper
;; under a synthetic one, so both functions survive.

(module
  ;; `$get` takes two arguments; the exported "get" takes one and calls it.
  (func $get (param $i i32) (param $j i32) (result i32)
    (i32.add (local.get $i) (local.get $j)))
  (func (export "get") (param $i i32) (result i32)
    (call $get (local.get $i) (i32.const 100)))

  ;; The same shape with no arguments at all, which is where the recursion is
  ;; unconditional.
  (func $tick (result i32) (i32.const 7))
  (func (export "tick") (result i32)
    (i32.mul (call $tick) (i32.const 2)))

  ;; Two exports on one unnamed function, the FIRST of which collides.
  (func $both (result i32) (i32.const 1))
  (func (export "both") (export "both_alias") (result i32)
    (i32.add (call $both) (i32.const 10)))

  ;; A collision where the `$id` is declared AFTER the exported wrapper — a
  ;; single-pass scan cannot see it yet, so the ids are collected first.
  (func (export "later") (result i32)
    (i32.add (call $later) (i32.const 1000)))
  (func $later (result i32) (i32.const 5))

  ;; And the ordinary case, unchanged: an unnamed exported function with no
  ;; competing id is still reached by its export name.
  (func (export "plain") (result i32) (i32.const 42))
)

(assert_return (invoke "get" (i32.const 5)) (i32.const 105))
(assert_return (invoke "tick") (i32.const 14))
(assert_return (invoke "both") (i32.const 11))
(assert_return (invoke "both_alias") (i32.const 11))
(assert_return (invoke "later") (i32.const 1005))
(assert_return (invoke "plain") (i32.const 42))
