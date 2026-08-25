;; vybe-test: wast/wat_component/test_a_resource_handle_binds_by_name
;; hand-written against proposals/component-model/design/mvp/Explainer.md
;;   §Type definitions — `(type $t (resource (rep i32)))` — and
;;   CanonicalABI.md §Flattening (:3198): `OwnType() | BorrowType() : ['i32']`.
;;
;; ▶▶ A HANDLE BINDS BY NAME; THE SOURCE WRITES AN INDEX. `component::ValType`
;; spells an owned handle `Own(String)` — it binds to whatever is registered
;; under that name — while `(own $file)` is a TYPEIDX. `TypeDecl::Resource`
;; carries the `(type $id …)` binder, and it is the only record connecting the
;; two. Without it every `(own …)` was unresolvable and refused.
;;
;; ⛔ AN INVENTED NAME WOULD BE WORSE THAN A REFUSAL. A wrong name still binds
;; — to a DIFFERENT resource — so the component links, runs, and hands the
;; callee someone else's object. That is why an UNNAMED `(type (resource …))`
;; refuses rather than being given a positional stand-in, and why a typeidx
;; that names a non-resource gets a different message from one out of range:
;; they are different mistakes.
;;
;; A handle flattens to ONE core `i32` — the index into the instance's handle
;; table, not the representation. The callee here doubles it, so 21 → 42 proves
;; the handle crossed as a scalar rather than being lifted as a struct or
;; dropped for want of a name.

(component
  (core module $m
    (func (export "twice") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "twice" (core func $d))

  (type $file (resource (rep i32)))
  (type $ft (func (param "h" (own $file)) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))
  (canon lower (func $lifted) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 21))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
