;; vybe-test: wast/wat_programs/power_of_two
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $pow2 (export "pow2") (param $n i32) (result i32)
    i32.const 1
    local.get $n
    i32.shl)
)
