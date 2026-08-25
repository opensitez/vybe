;; vybe-test: wast/wat_component/test_a_bare_name_collides_with_an_annotated_member
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2829
;;   §Name uniqueness, clause 2:
;;     * If one name is `l` and the other name is `[*]l.l` … they *are not*
;;       strongly-unique.
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   export: "bar" is not strongly-unique against "[method]foo.bar"
;;   (Explainer.md:2829 — a bare `bar` and an annotated `….bar` generate the
;;   same binding name)
;;
;; ▶▶ CLAUSE 2'S NOTATION IS AMBIGUOUS AND THE EXAMPLES SETTLE IT. Written
;; `[*]l.l`, it reads as though BOTH dotted labels must equal the bare name.
;; They do not. The spec's own lists are decisive:
;;
;;   legal together:  foo, foo-bar, [constructor]foo, [method]foo.bar,
;;                    [method]foo.baz, foo:bar/baz
;;   adding any is an error:  foo, foo-BAR, [constructor]foo-BAR,
;;                    [method]foo.foo, [method]foo.BAR, foo:bar/baz, bar
;;
;; `bar` is an error, and NOTHING in the legal set collides with `bar` under
;; clause 3 — `foo-bar`, `foo.bar` and `foo:bar/baz` are all unequal to `bar`.
;; Only `[method]foo.bar` can explain it, and only if the SECOND label is what
;; matters. Meanwhile `[method]foo.bar` sits happily beside `foo`, so the FIRST
;; label does not.
;;
;; The stated rationale confirms the reading: the rule exists for "pathological
;; cases where two unique-in-the-component names get mapped to the same
;; source-language identifier". A method `bar` and a free function `bar` both
;; become `bar` in a generated binding.

(component
  (core module $m
    (func (export "id") (param i32) (result i32) (local.get 0)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "id" (core func $d))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $d) (func $lifted (type $ft)))

  (export "[method]foo.bar" (func $lifted))
  (export "bar" (func $lifted))
)
