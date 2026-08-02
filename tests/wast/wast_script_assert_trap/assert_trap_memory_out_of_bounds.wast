;; vybe-test: wast/wast_script_assert_trap/assert_trap_memory_out_of_bounds
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (memory 1) (func (export "f") (result i32) i32.const 65536 i32.load))
(assert_trap (invoke "f") "out of bounds memory access")
