;; vybe-test: wast/wat_component/test_with_clause_addresses_a_canon_row
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §3
;;   `core:instantiatearg ::= (with <core:name> (instance <core:inlineexport>*))`
;;   `core:inlineexport   ::= (export <core:name> <core:externidx>)`
;;
;; ⛔ THIS IS THE TEST THAT `@N` IS NO LONGER THE ADDRESSING MECHANISM.
;;
;; Every other canon test in this tree reaches a built-in by writing the
;; canonidx into the import name — `(import "canon" "thread.spawn-ref@0" …)`.
;; That was never how the Component Model addresses a canonical definition; it
;; was the only spelling a module-only front end could express.
;;
;; Here the core module imports a slot it names "spawn", which is not a
;; built-in's name and carries no index. The `with` clause is what fills that
;; slot, naming the canon row by its `(core func $sr)` BINDER. **There is no
;; `@N` anywhere in this file.**
;;
;; The trap message is the discriminator and MUST be read — `run-fail` is green
;; on any failure, and the null funcref would fail on its own:
;;
;;   canon thread.spawn-ref: `shared` requests a PREEMPTIVE thread ... (trap)
;;
;; It reports `thread.spawn-ref`, which the source never spells at the import,
;; and it honours `shared`, which only that canon row declares. Both facts
;; travelled through the binder. Its control is
;; `test_with_clause_is_required_to_reach_a_row`, which drops the `with` clause
;; and gets a completely different failure.

(component
  (canon thread.spawn-ref shared (core type 0) (core func $sr))
  (core module $m
    (import "canon" "spawn"
      (func $spawn (param funcref) (param i32) (result i32)))
    (func (export "_start")
      ref.null func
      i32.const 0
      call $spawn
      drop))
  (core instance (instantiate $m
    (with "canon" (instance (export "spawn" (func $sr))))))
)
