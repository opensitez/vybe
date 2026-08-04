;; vybe-test: wast/wat_execution/global_mut_increment
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (global $g (mut i32) (i32.const 0))
  (func (export "_start")
    global.get $g
    i32.const 1
    i32.add
    global.set $g
    global.get $g
    i32.const 1 call $vybe_check_i32
    global.get $g
    i32.const 1
    i32.add
    global.set $g
    global.get $g
    i32.const 2 call $vybe_check_i32))
