;; vybe-test: wast/wat_algorithms/test_triangular_number
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
        (func $tri (param $n i32) (result i32)
          local.get $n local.get $n i32.const 1 i32.add i32.mul i32.const 2 i32.div_u)
        (func (export "_start") i32.const 100 call $tri i32.const 5050 call $vybe_check_i32))
