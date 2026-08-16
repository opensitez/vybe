;; vybe-test: wast/wast_script/assert_malformed_quote
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: run
;;
;; `(module quote "…")` defers the TEXT, and `assert_malformed` says that text
;; must fail to PARSE.
;;
;; The fixture here used to quote `(module (func (result i32)))` and expect
;; "unexpected token". That module parses perfectly — it is *invalid*, not
;; malformed (wasmtime: "type mismatch: expected i32 but nothing on stack",
;; from the validator, after parsing succeeded). The two are different
;; assertions and only one of them is about the text. It went unnoticed because
;; the directive lowered to nothing at all.

;; Unbalanced parentheses — the text cannot be read at all.
(assert_malformed (module quote "(module (func (result i32)") "unexpected end")
(assert_malformed (module quote "(module))") "unexpected token")

;; A quoted module that IS well-formed must not be asserted malformed, so the
;; well-formed counterpart is checked the ordinary way: it parses, compiles and
;; runs. Without this the file would pass under "reject everything".
(module quote "(module (func (export \"seven\") (result i32) (i32.const 7)))")
(assert_return (invoke "seven") (i32.const 7))
