;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_alignment
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (memory 1) (func i32.const 0 i32.load align=8 drop)) "alignment must not be larger than natural")
