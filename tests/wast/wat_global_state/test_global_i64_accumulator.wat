;; vybe-test: wast/wat_global_state/test_global_i64_accumulator
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
        (global $acc (mut i64) (i64.const 0))
        (func $add (param $n i64) global.get $acc local.get $n i64.add global.set $acc)
        (func (export "_start")
          i64.const 1000000000 call $add i64.const 2000000000 call $add
          global.get $acc i64.const 3000000000 call $vybe_check_i64))
