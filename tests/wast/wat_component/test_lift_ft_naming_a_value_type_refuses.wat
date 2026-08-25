;; vybe-test: wast/wat_component/test_lift_ft_naming_a_value_type_refuses
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §8
;;   — a component has ONE type index space holding both value and function
;;   types; the VM keeps two tables (`canon_types`, `canon_functypes`).
;;
;; ⛔ THIS IS THE POSITIONAL-ALIGNMENT PROOF, and it is the whole reason both
;; VM tables are `Vec<Option<_>>` rather than dense vectors of what they hold.
;;
;; The source type space here is:
;;
;;   typeidx 0 → `u32`                  (a VALUE type)
;;   typeidx 1 → `(func (result u32))`  (a FUNCTION type)
;;
;; `canon lift` names typeidx 0 as its `$ft`, which is wrong, and must be told
;; so. Under DENSE tables the single function type would sit at functype index
;; 0, so lifting `(type $v)` would find it and SUCCEED — silently lifting with
;; a signature the source never named. That is the `GLOBAL_GET` defect: one
;; integer meaning two things depending on which table reads it.
;;
;; `run-fail` is green on any failure, so the message MUST be read:
;;
;;   canon lift: $ft 0 is a declared typeidx but does not hold a FUNCTION type
;;
;; "declared but wrong kind" and "out of range" are deliberately different
;; messages, because they are different mistakes: a stale index versus naming
;; the wrong kind of type.

(component
  (type $v u32)
  (type $ft (func (result u32)))
  (canon lift (core func 999) (func $lifted (type $v)))
  (core module $c
    (import "canon" "lift@0" (func $l (result i32)))
    (func (export "_start")
      (drop (call $l))))
  (core instance (instantiate $c))
)
