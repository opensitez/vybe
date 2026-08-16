;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_invalid_utf8
;; origin: languages/wast/tests/wast/test_wast_script_assert_malformed.rs
;; vybe-test-mode: run
;;
;; Invalid UTF-8 is malformed in a NAME — an import or export name is a
;; character string and must decode. It is NOT malformed in a data segment:
;; a data string is a BYTE string, and `\ff` there is an ordinary byte.
;;
;; This fixture used to quote `(module (data "\ff"))` and expect
;; "invalid utf-8 encoding". wasmtime accepts that module without complaint —
;; the assertion was simply false, and it went unnoticed because the directive
;; lowered to nothing.

(assert_malformed
  (module quote "(module (func (export \"\\ff\")))")
  "malformed UTF-8 encoding")

;; The byte string that is NOT malformed, checked the ordinary way — otherwise
;; "reject every quoted module" would satisfy this file.
(module quote "(module (memory 1) (data (i32.const 0) \"\\ff\") (func (export \"first_byte\") (result i32) (i32.load8_u (i32.const 0))))")
(assert_return (invoke "first_byte") (i32.const 255))
