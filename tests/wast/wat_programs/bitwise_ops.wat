;; vybe-test: wast/wat_programs/bitwise_ops
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $flags (export "flags") (param $a i32) (param $b i32) (result i32)
    local.get $a
    local.get $b
    i32.and
    local.get $a
    local.get $b
    i32.or
    i32.xor)
)
