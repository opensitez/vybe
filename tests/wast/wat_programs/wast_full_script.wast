;; vybe-test: wast/wat_programs/wast_full_script
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $add (export "add") (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.add)
  (func $sub (export "sub") (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.sub)
  (func $mul (export "mul") (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.mul)
)
(assert_return (invoke "add" (i32.const 10) (i32.const 5)) (i32.const 15))
(assert_return (invoke "sub" (i32.const 10) (i32.const 5)) (i32.const 5))
(assert_return (invoke "mul" (i32.const 10) (i32.const 5)) (i32.const 50))
(assert_return (invoke "add" (i32.const 0) (i32.const 0)) (i32.const 0))
(assert_return (invoke "add" (i32.const -1) (i32.const 1)) (i32.const 0))
