;; vybe-test: wast/wat_component/test_an_unnamed_resource_cannot_be_owned
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md
;;   §Type definitions, and `component::ValType::Own(String)`.
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   component type 1: param "h": `(own 0)` names an UNNAMED resource type; a
;;   handle binds by name, so the `(type $id (resource …))` it points at has to
;;   have a binder
;;
;; ▶▶ THE REFUSAL IS THE POINT, AND SO IS ITS SPECIFICITY. `ValType::Own` holds
;; a NAME and binds to whatever is registered under it. An unnamed resource has
;; no name to bind to, and any stand-in — the typeidx as a string, a generated
;; label — would bind the handle to a DIFFERENT resource. That does not fail
;; loudly: it links, it runs, and the callee receives someone else's object.
;;
;; The type still OCCUPIES index 0. Skipping an unnamed declaration instead
;; would renumber every later typeidx, which is the same positional hazard the
;; `Vec<Option<_>>` type tables exist to avoid.
;;
;; Three refusals are deliberately distinct, because they are three different
;; mistakes and a reader needs to know which one they made:
;;
;;   unnamed resource      → this file
;;   typeidx is not a resource  → "declared but is not a resource type"
;;   typeidx out of range       → "is not in the component type space (have N)"

(component
  (type (resource (rep i32)))
  (type $ft (func (param "h" (own 0)) (result u32)))
  (canon lift (core func 999) (func $lifted (type $ft)))
)
