;; vybe-test: wast/wat_component/test_a_constructor_name_is_strongly_unique_with_its_label
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2827
;;   §Name uniqueness, clause 1:
;;     * If one name is `l` and the other name is `[constructor]l` (for the same
;;       `label` `l`), they *are* strongly-unique.
;;
;; ▶▶ THE ONE CASE THE FOLD MUST *NOT* CATCH. `foo` and `[constructor]foo`
;; strip and fold to the SAME key — `foo` — so clause 3 alone would reject
;; them. Clause 1 exempts the pair, and the spec's own legal set contains
;; exactly it:
;;
;;   foo, foo-bar, [constructor]foo, [method]foo.bar, [method]foo.baz, foo:bar/baz
;;
;; A folded-key map cannot express this on its own, which is the whole reason
;; the table keeps the name AS WRITTEN as its value: the exemption is decided by
;; comparing the RAW names, not their keys.
;;
;; ⛔ AND IT MUST COMPARE THEM RAW. `[constructor]foo-BAR` against `foo-bar` is
;; listed as an ERROR, even though `foo-BAR` and `foo-bar` are the same label
;; after folding. So the exemption is `[constructor]` + the identical raw label,
;; not the identical key — checking it after folding would let that pair
;; through. `test_export_names_must_be_strongly_unique` pins the folding half.
;;
;; This file also proves the component still RUNS with both names exported: the
;; lowered function answers 21 * 2 = 42, so a name check that refused a legal
;; pair would show up as a compile error rather than a wrong number.

(component
  (core module $m
    (func (export "double") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "double" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))

  (export "foo" (func $lifted))
  (export "[constructor]foo" (func $lifted))
  (export "foo-bar" (func $lifted))
  (export "[method]foo.bar" (func $lifted))
  (export "[method]foo.baz" (func $lifted))

  (canon lower (func $lifted) (core func $lo))
  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 21))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
