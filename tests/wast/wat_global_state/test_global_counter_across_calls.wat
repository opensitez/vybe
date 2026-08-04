;; vybe-test: wast/wat_global_state/test_global_counter_across_calls
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
        (global $c (mut i32) (i32.const 0))
        (func $inc global.get $c i32.const 1 i32.add global.set $c)
        (func (export "_start")
          call $inc call $inc call $inc call $inc global.get $c i32.const 4 call $vybe_check_i32))
