;; vybe-test: wast/wat_component/test_a_nested_component_runs_when_instantiated
;; hand-written against proposals/component-model/design/mvp/Explainer.md
;;   §Instance definitions — `(instance (instantiate <componentidx> <arg>*))`
;;   and §889 — a component type carries TWO named lists, imports and exports.
;;
;; ▶▶ A NESTED COMPONENT IS DECLARED, NOT WALKED — and this is where it runs.
;;
;; It used to be walked INLINE, which was wrong twice over: its core modules
;; executed where the component was written, whether or not anything
;; instantiated it, and a component instantiated TWICE would still only ever
;; have run once. `(core module …)` had already solved exactly this shape by
;; declaring into `CoreModules` and running at `(core instance …)`; this is the
;; same treatment one level up.
;;
;; The full chain, and every link is a separate producer:
;;
;;   (component $inner …)          declared into the COMPONENT index space
;;   (instance $i (instantiate $inner))   RUNS it; its exports become the
;;                                        INSTANCE's export table
;;   (alias export $i "double" …)  reaches into that table
;;   (canon lower (func $reached)) calls it
;;
;; 21 × 2 = 42, so a chain that reached nothing returns 21 and one that dropped
;; the argument returns 0.
;;
;; ⛔ THE EXPORT TABLE CROSSES A SCOPE BOUNDARY AND THE INDICES MUST STILL BE
;; VALID. `walk_component` restores the enclosing component's NAME maps on the
;; way out, but the funcidx values it returns index `comp_func_space`, which is
;; AST PAYLOAD and deliberately shared across nesting. Scoping that space would
;; make every returned index dangle — the inner's `canon lift` would land in a
;; vector that is thrown away.
;;
;; ⛔ Deleting the `(instance …)` line makes this refuse with `$i is not bound
;; in the component instance index space`, which is the proof that the
;; instantiation — not the declaration — is what produced the instance.

(component
  (component $inner
    (core module $m
      (func (export "double") (param i32) (result i32)
        (i32.mul (local.get 0) (i32.const 2))))
    (core instance $mi (instantiate $m))
    (alias core export $mi "double" (core func $d))
    (type $ft (func (param "a" u32) (result u32)))
    (canon lift (core func $d) (func $lifted (type $ft)))
    (export "double" (func $lifted))
  )

  (instance $i (instantiate $inner))
  (alias export $i "double" (func $reached))
  (canon lower (func $reached) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 21))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
