;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_binary_invalid_section
;; origin: languages/wast/tests/wast/test_wast_script_assert_malformed.rs
;; vybe-test-mode: compile

(assert_malformed (module binary "\00asm\01\00\00\00\ff\00") "malformed section id")
