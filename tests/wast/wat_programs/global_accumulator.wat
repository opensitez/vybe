;; vybe-test: wast/wat_programs/global_accumulator
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (global $total (mut i32) (i32.const 0))
  (func $add (export "add") (param $n i32)
    global.get $total
    local.get $n
    i32.add
    global.set $total)
  (func $get (export "get") (result i32)
    global.get $total)
  (func $reset (export "reset")
    i32.const 0
    global.set $total)
)
