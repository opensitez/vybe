;; vybe-test: wast/wast_script_assert_trap/assert_trap_table_out_of_bounds
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (table 1 funcref) (func (export "f") (result funcref) i32.const 1 table.get 0))
(assert_trap (invoke "f") "out of bounds table access")
