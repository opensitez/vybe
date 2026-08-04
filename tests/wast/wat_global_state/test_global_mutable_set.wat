;; vybe-test: wast/wat_global_state/test_global_mutable_set
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
        (global $g (mut i32) (i32.const 0))
        (func (export "_start") i32.const 77 global.set $g global.get $g i32.const 77 call $vybe_check_i32))
