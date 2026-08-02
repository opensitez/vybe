;; vybe-test: wast/wast_script/assert_malformed_binary
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(assert_malformed (module binary "\00asm") "magic header not detected")
