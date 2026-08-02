;; vybe-test: wast/wast_script/assert_trap_oob_memory
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (memory 1) (func (export "load") (param i32) (result i32) local.get 0 i32.load))
(assert_trap (invoke "load" (i32.const 65536)) "out of bounds memory access")
