;; vybe-test: wast/wat_try_table/tag_declaration_forms
;; vybe-test-mode: run
;;
;; A tagtype is a TYPEUSE — `(type $t)`, an inline `(param …)*` / `(result …)*`,
;; or both — not just a param list. Our grammar spelled it
;; `("(" "func" param* ")") | param*`, which parses neither of the two forms the
;; spec's own `tag.wast` opens with:
;;
;;   (tag (result i32))        ;; tag.wast:19 — the text `assert_invalid` quotes
;;   (tag (type $t1))          ;; tag.wast:35
;;
;; The first matters even though it is INVALID: a validation assertion can only
;; reject text the parser can read, so `tag.wast` died at its own fixture rather
;; than at the rule it was written to test. The whole file was unreachable, and
;; with it every tag-declaration form below.
;;
;; What is exercised here is the ARITY each form yields, because that is what
;; `catch` matches against and what a wrong reading silently gets wrong: a
;; `(type $t)` tag whose arity came back 0 delivers no payload to its handler.
;;
;; Spec-format so `wasmtime wast` arbitrates.

(module
  (type $unary (func (param i32)))
  (type $binary (func (param i32 i32)))
  (type $nullary (func))

  ;; Every declaration form the spec allows, side by side.
  (tag $anon)                              ;; no params at all
  (tag $inline (param i32))                ;; inline typeuse
  (tag $inline2 (param i32 i32))           ;; multi-param inline
  (tag $viatype (type $unary))             ;; by type reference
  (tag $viatype2 (type $binary))
  (tag $viatype0 (type $nullary))
  (tag $both (type $unary) (param i32))    ;; reference AND inline restatement

  ;; ── A zero-arity tag delivers nothing ────────────────────────────────
  (func (export "anon_roundtrip") (result i32)
    (block $h (result i32)
      (block $h2
        (try_table (catch $anon $h2)
          (throw $anon))
        (return (i32.const -1)))
      (i32.const 1)))

  ;; ── Inline params arrive as the handler's values ─────────────────────
  (func (export "inline_payload") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $inline $h)
        (throw $inline (i32.const 7)))))

  (func (export "inline2_payload") (result i32)
    (block $h (result i32 i32)
      (try_table (result i32 i32) (catch $inline2 $h)
        (throw $inline2 (i32.const 3) (i32.const 4))))
    i32.add)

  ;; ── A `(type $t)` tag has the referenced type's arity ────────────────
  ;; Read as arity 0 the payload never reaches the handler.
  (func (export "viatype_payload") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $viatype $h)
        (throw $viatype (i32.const 9)))))

  (func (export "viatype2_payload") (result i32)
    (block $h (result i32 i32)
      (try_table (result i32 i32) (catch $viatype2 $h)
        (throw $viatype2 (i32.const 20) (i32.const 22))))
    i32.add)

  (func (export "viatype0_roundtrip") (result i32)
    (block $h2
      (try_table (catch $viatype0 $h2)
        (throw $viatype0))
      (return (i32.const -1)))
    (i32.const 5))

  ;; ── Reference plus inline restatement is still arity 1 ───────────────
  (func (export "both_payload") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $both $h)
        (throw $both (i32.const 12)))))

  ;; ── Identity is the DECLARATION, not the type ────────────────────────
  ;; `$inline` and `$viatype` have the same signature and must never match
  ;; each other.
  (func (export "same_signature_never_matches") (result i32)
    (block $h (result i32)
      (try_table (result i32) (catch $viatype $h)
        (throw $inline (i32.const 1)))))
)

(assert_return (invoke "anon_roundtrip") (i32.const 1))
(assert_return (invoke "inline_payload") (i32.const 7))
(assert_return (invoke "inline2_payload") (i32.const 7))
(assert_return (invoke "viatype_payload") (i32.const 9))
(assert_return (invoke "viatype2_payload") (i32.const 42))
(assert_return (invoke "viatype0_roundtrip") (i32.const 5))
(assert_return (invoke "both_payload") (i32.const 12))
(assert_exception (invoke "same_signature_never_matches"))
