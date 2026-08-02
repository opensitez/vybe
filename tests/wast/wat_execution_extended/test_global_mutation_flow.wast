;; vybe-test: wast/wat_execution_extended/test_global_mutation_flow
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs
;; vybe-test-mode: compile

(module
  (global $g (mut i32) (i32.const 10))
  (func (export "inc") (result i32)
    global.get $g
    i32.const 5
    i32.add
    global.set $g
    global.get $g))
(assert_return (invoke "inc") (i32.const 15))
(assert_return (invoke "inc") (i32.const 20))
