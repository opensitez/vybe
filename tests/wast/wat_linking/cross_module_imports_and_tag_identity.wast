;; vybe-test: wast/wat_linking/cross_module_imports_and_tag_identity
;; origin: the `(register "m")` / `(import "m" "e")` wiring, which lowered to
;; two inert marker calls — `__wasm_import` and `__wasm_register` appeared
;; NOWHERE outside the walker that emitted them.
;; vybe-test-mode: run
;;
;; Every module of a wast script becomes a class whose exports are its static
;; methods. An import is a second NAME for one of those methods, so linking is
;; a name resolution: `(register "m")` publishes the class, and the importing
;; module's alias must resolve against THAT class. It did not — the alias was
;; resolved against the importing module's own class, where a same-named stub
;; sat with no body, so every cross-module call returned null.
;;
;; That was invisible because `wast_script_register_get/register_and_import`
;; was `vybe-test-mode: compile`: it proved the script COMPILED and never ran
;; the assertion that would have caught it.
;;
;; The exception cases matter beyond linking. A tag's identity is its
;; DECLARATION, and an import names someone else's declaration — so:
;;
;;   * a tag imported under two different local aliases is ONE entity;
;;   * an exception thrown by the exporting module must match a `catch` on the
;;     imported alias in the importing module;
;;   * and it must NOT match a same-signature tag the importer declared itself,
;;     which is exactly the case a name-keyed entity table gets wrong.
;;
;; Spec-format so `wasmtime wast` arbitrates every one of these.

(module
  (tag $e0 (export "e0") (param i32))
  (tag $other (export "other") (param i32))
  (func (export "five") (result i32) (i32.const 5))
  (func (export "add") (param i32 i32) (result i32)
    (i32.add (local.get 0) (local.get 1)))
  (func (export "throw-e0") (param i32) (throw $e0 (local.get 0)))
  (func (export "throw-other") (param i32) (throw $other (local.get 0)))
)

(register "lib")

(module
  (func $five (import "lib" "five") (result i32))
  (func $add (import "lib" "add") (param i32 i32) (result i32))
  (func $throw-e0 (import "lib" "throw-e0") (param i32))
  (func $throw-other (import "lib" "throw-other") (param i32))

  ;; The SAME exported tag, imported twice under different local names.
  (tag $imported (import "lib" "e0") (param i32))
  (tag $alias (import "lib" "e0") (param i32))
  ;; A tag this module declares itself, with the SAME signature as $imported.
  ;; A fresh declaration is a fresh entity — it must never match $imported.
  (tag $mine (param i32))

  ;; ── plain cross-module calls ─────────────────────────────────────────
  (func (export "call_no_args") (result i32)
    (call $five))

  (func (export "call_with_args") (result i32)
    (call $add (i32.const 40) (i32.const 2)))

  ;; A tail call across modules reuses the frame the same way.
  (func (export "tail_call_across_modules") (result i32)
    (return_call $five))

  ;; ── an imported tag is the EXPORTING module's entity ─────────────────
  (func (export "catch_imported") (result i32)
    (block $h (result i32)
      (try_table (catch $imported $h)
        (call $throw-e0 (i32.const 11)))
      (i32.const -1)))

  ;; Two aliases of one import are one entity: throw arrives on $imported,
  ;; the clause names $alias, and it still matches.
  (func (export "catch_through_alias") (result i32)
    (block $h (result i32)
      (try_table (catch $alias $h)
        (call $throw-e0 (i32.const 21)))
      (i32.const -1)))

  ;; ── identity, not signature ──────────────────────────────────────────
  ;; $mine has the same signature as $imported and must not catch it; the
  ;; enclosing catch_all is what actually fires.
  (func (export "own_tag_does_not_catch_imported") (result i32)
    (block $outer
      (block $wrong (result i32)
        (try_table (catch_all $outer)
          (try_table (catch $mine $wrong)
            (call $throw-e0 (i32.const 31))))
        (unreachable))
      (drop))
    (i32.const 41))

  ;; ...and a different EXPORTED tag is a different entity too.
  (func (export "other_exported_tag_does_not_match") (result i32)
    (block $outer
      (block $wrong (result i32)
        (try_table (catch_all $outer)
          (try_table (catch $imported $wrong)
            (call $throw-other (i32.const 51))))
        (unreachable))
      (drop))
    (i32.const 61))

  ;; An exception raised in another module and caught by nothing here still
  ;; unwinds out of this one.
  (func (export "uncaught_crosses_back") (param i32)
    (call $throw-e0 (local.get 0)))
)

;; ── linking reaches the exporting module's code ─────────────────────────
(assert_return (invoke "call_no_args") (i32.const 5))
(assert_return (invoke "call_with_args") (i32.const 42))
(assert_return (invoke "tail_call_across_modules") (i32.const 5))

;; ── an imported tag is one entity with the exporter's ───────────────────
(assert_return (invoke "catch_imported") (i32.const 11))
(assert_return (invoke "catch_through_alias") (i32.const 21))

;; ── and is told apart from every other tag by DECLARATION ───────────────
(assert_return (invoke "own_tag_does_not_catch_imported") (i32.const 41))
(assert_return (invoke "other_exported_tag_does_not_match") (i32.const 61))
(assert_exception (invoke "uncaught_crosses_back" (i32.const 71)))
