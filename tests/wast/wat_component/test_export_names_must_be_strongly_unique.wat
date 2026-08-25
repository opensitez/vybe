;; vybe-test: wast/wat_component/test_export_names_must_be_strongly_unique
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2826
;;   §Name uniqueness, clause 3:
;;     * Lowercase all the `acronym`s (uppercase letters) in both names.
;;     * Strip any `[...]` annotation prefix from both names.
;;     * The names are strongly-unique if the resulting strings are unequal.
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   export: "foo-BAR" is not strongly-unique against "foo-bar"
;;   (Explainer.md:2826 — annotations stripped and acronyms lowercased, both
;;   are `foo-bar`)
;;
;; ▶▶ EXACT-STRING EQUALITY IS NOT THE RULE. `foo-bar` and `foo-BAR` are two
;; different strings and the same NAME, because a `label` fragment is either
;; entirely lower-case (`word`) or entirely upper-case (`acronym`) — so
;; lowercasing the acronyms is lowercasing the string. The spec lists exactly
;; this pair as a validation error.
;;
;; ⛔ THE FOLD IS ONLY THE COLLISION KEY. The table stores the name AS WRITTEN
;; and folds only what it compares on, because an export name is an identifier
;; a host matches against byte-for-byte. Folding the stored name would be the
;; namespace tree's lowercase-canonical mistake in a second place.
;;
;; Both exports name the SAME function on purpose: exporting one function under
;; two names is legal, so the only thing under test here is the name rule.

(component
  (core module $m
    (func (export "id") (param i32) (result i32) (local.get 0)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "id" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))

  (export "foo-bar" (func $lifted))
  (export "foo-BAR" (func $lifted))
)
