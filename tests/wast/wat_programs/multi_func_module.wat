;; vybe-test: wast/wat_programs/multi_func_module
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $double (param $x i32) (result i32)
    local.get $x i32.const 2 i32.mul)
  (func $triple (param $x i32) (result i32)
    local.get $x i32.const 3 i32.mul)
  (func $sextuple (export "sextuple") (param $x i32) (result i32)
    local.get $x call $double call $triple)
)
