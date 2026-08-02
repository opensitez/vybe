;; vybe-test: wast/wat_lexical/id_special_chars
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile-fail

(module (func $func-name!@#%^&*()_+{}|:<>?-=[]\;',./))
