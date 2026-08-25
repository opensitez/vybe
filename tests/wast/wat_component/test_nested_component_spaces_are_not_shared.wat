;; vybe-test: wast/wat_component/test_nested_component_spaces_are_not_shared
;; vybe-test-mode: compile-fail
;; hand-written against proposals/component-model/design/mvp/Explainer.md §Components
;;
;; ▶▶ AN INNER COMPONENT'S BINDINGS DO NOT ESCAPE IT — **EVEN AFTER IT RUNS**.
;;
;; The inner component is INSTANTIATED here, on purpose. That is the whole
;; difference between this file and the version it replaces: previously the
;; inner was walked inline, so `$hidden` really was created and really was
;; scoped out. Once nested components became DECLARED rather than walked, the
;; old file kept passing — but vacuously, because `$hidden` was never created
;; at all. Same message, weaker reason.
;;
;; Its own header said what to do about that: *"If it ever refuses for a
;; different reason, this test has stopped testing scoping."* So the
;; instantiation is here to make the binding genuinely exist before the outer
;; component reaches for it.
;;
;; ⛔ THE BUG THIS GUARDS WAS SILENT. `walk_component` once scoped only its core
;; MODULE list and left the core function, core instance and both type spaces on
;; the walker, shared with every nested component — so an inner `(type …)`
;; renumbered the outer's type space and an inner alias stayed visible outside
;; it. Every one of those is a small integer, so nothing ever looked wrong; a
;; call simply reached a different function than the source named.
;;
;; `compile-fail` is green on ANY compile failure, so the message MUST be read:
;;
;;   canon: `$hidden` is not bound in the core func index space
;;
;; ⛔ If it ever refuses for a different reason — a parse error, an unknown
;; instantiation target, a missing export — this test has stopped testing
;; scoping and must be repaired rather than re-baselined.

(component
  (component $inner
    (core module $lib
      (func (export "answer") (result i32)
        (i32.const 1)))
    (core instance $li (instantiate $lib))
    ;; Bound in the INNER component's core function index space.
    (alias core export $li "answer" (core func $hidden))
  )

  ;; It RUNS. `$hidden` is genuinely created — and still must not be visible
  ;; out here.
  (instance (instantiate $inner))

  (core module $main
    (import "l" "f" (func $f (result i32)))
    (func (export "get") (result i32)
      (call $f)))
  (core instance (instantiate $main
    (with "l" (instance (export "f" (func $hidden))))))
)
