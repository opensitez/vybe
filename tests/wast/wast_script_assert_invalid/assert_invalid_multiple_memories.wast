;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_multiple_memories
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (memory 1) (memory 1)) "multiple memories")
