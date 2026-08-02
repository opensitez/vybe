;; vybe-test: wast/wat_instructions/compile_global_counter
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module
  (global $count (mut i32) (i32.const 0))
  (func $inc (export "inc")
    global.get $count
    i32.const 1
    i32.add
    global.set $count)
  (func $get (export "get") (result i32)
    global.get $count)
)
