;; vybe-test: wast/wat_global_state/test_global_immutable_read
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (global $g i32 (i32.const 314))
        (func (export "_start") global.get $g i32.const 314 call $vybe_check_i32))
