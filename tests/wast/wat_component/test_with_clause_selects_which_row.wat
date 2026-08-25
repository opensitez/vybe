;; vybe-test: wast/wat_component/test_with_clause_selects_which_row
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §3
;;   `core:inlineexport ::= (export <core:name> <core:externidx>)`
;;
;; Two canon rows of the SAME built-in, differing only in `shared?`:
;;
;;   canonidx 0 — `thread.spawn-ref`          (plain)
;;   canonidx 1 — `thread.spawn-ref shared`
;;
;; The core module imports one slot called "spawn". Which definition it reaches
;; is decided by the `with` clause naming a BINDER — here `$shared`, canonidx 1.
;;
;; The trap message is the discriminator and MUST be read. `run-fail` is green
;; on any failure and the null funcref would fail on its own:
;;
;;   canon thread.spawn-ref: `shared` requests a PREEMPTIVE thread ... (trap)
;;
;; If the wiring ignored the binder and took the first row — or fell back to the
;; import name — this would report `null funcref` instead, exactly as the same
;; component wired to `$plain` does. Two rows of one built-in is what makes that
;; distinguishable: the built-in's NAME cannot tell them apart, only the index
;; the binder carries can.

(component
  (canon thread.spawn-ref (core type 0) (core func $plain))
  (canon thread.spawn-ref shared (core type 0) (core func $shared))
  (core module $m
    (import "canon" "spawn"
      (func $s (param funcref) (param i32) (result i32)))
    (func (export "_start")
      ref.null func
      i32.const 0
      call $s
      drop))
  (core instance (instantiate $m
    (with "canon" (instance (export "spawn" (func $shared))))))
)
