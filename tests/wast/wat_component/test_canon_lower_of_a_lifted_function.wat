;; vybe-test: wast/wat_component/test_canon_lower_of_a_lifted_function
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon lift (:3632), §canon lower (:3855), and Binary.md:297
;;     0x01 0x00 f:<funcidx> opts:<opts> => (canon lower f opts (core func))
;;
;; ▶▶ THE LIFT ↔ LOWER ROUND TRIP — the centre of the Canonical ABI.
;;
;;   core 21  →  canon lower  →  canon lift  →  core `double`  →  42
;;
;; `lower` lifts the flat core args into component values and calls the
;; component function; that function IS the `canon lift`, which lowers them
;; back to flat and calls the core callee. Both directions of the ABI run, on
;; one value, in one call.
;;
;; ⛔ THIS DID NOT NEED THE LINKER, contrary to what cmplan.md claimed for
;; months. `canon_lower(callee, ft, opts, flat_args)` takes `ft` as a
;; PARAMETER supplied by the callee, and the row above carries NO `ft`
;; immediate at all. For a function lifted in the SAME component that type is
;; the lift row's own `$ft`. What was actually missing was the COMPONENT
;; FUNCTION INDEX SPACE — the same gap as the core function space, the type
;; space, and `VM::canon_defs` before them: `canon lift` DEFINES a component
;; function and nothing recorded it. The linker is only for component functions
;; that arrive as IMPORTS.
;;
;; 21 and 42 are chosen so the assert cannot pass by accident: the core callee
;; MULTIPLIES, so a round trip that dropped the argument would return 0 and one
;; that skipped the callee would return 21.
;;
;; There is no `@N` in this file — the `with` clause names the lowered core
;; function by its `(core func $lo)` binder.

(component
  (core module $m
    (func (export "double") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "double" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift  (core func $d) (func $lifted (type $ft)))
  (canon lower (func $lifted) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 21))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
