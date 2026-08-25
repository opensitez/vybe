;; vybe-test: wast/wat_component/test_component_functype_reaches_the_vm
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Binary.md
;;   §"Canonical Definitions" — `canon lift f opts ft:<typeidx>`.
;;
;; `VM::canon_functypes` had NO PRODUCER. It started `Vec::new()` and nothing
;; ever appended, so a component's `(type $ft (func …))` had nowhere to go and
;; `canon lift` trapped with `$ft 0 is not registered ... (have 0)` even when
;; the source declared the type. This is the carrier that fixes that: walker →
;; vybe_ast::canon::TypeDecl → compiler → merge at load, the same shape the
;; canon section already uses.
;;
;; The DISCRIMINATOR is which trap arrives. `run-fail` is green on any failure,
;; so the message MUST be read:
;;
;;   canon lift: $callee core funcidx 999 is out of range (have 4)
;;
;; `$callee` is checked AFTER `$ft` in `exec_canon_lift`, so reaching a callee
;; complaint at all proves `$ft` resolved. If this ever reports `$ft 0 is not
;; registered in VM::canon_functypes`, the carrier has stopped carrying.
;;
;; The callee is deliberately out of range rather than valid: a valid one is a
;; CHUNK index (see test_named_lift_callee_refuses_the_index_conflation), and
;; chunk 0 is the script chunk, so lifting it recurses until the stack dies.

(component
  (type $v u32)
  (type $ft (func (result u32)))
  (canon lift (core func 999) (func $lifted (type $ft)))
  (core module $c
    (import "canon" "lift@0" (func $l (result i32)))
    (func (export "_start")
      (drop (call $l))))
  (core instance (instantiate $c))
)
