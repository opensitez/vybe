;; vybe-test: wast/wat_algorithms/test_count_set_bits
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func (export "_start") i32.const 0xB7 i32.popcnt i32.const 6 call $vybe_check_i32))
