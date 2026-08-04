;; vybe-test: wast/wat_assignment/test_global_compound_update
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (global $g (mut i32) (i32.const 100))
        (func (export "_start")
          global.get $g i32.const 50 i32.sub global.set $g global.get $g i32.const 50 call $vybe_check_i32))
