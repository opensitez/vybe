;; vybe-test: wast/wat_recursion/test_tail_recursive_sum
;; origin: languages/wast/tests/wast/test_wat_recursion.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $sum (param $n i32) (param $acc i32) (result i32)
          local.get $n i32.eqz
          if (result i32) local.get $acc
          else local.get $n i32.const 1 i32.sub
               local.get $acc local.get $n i32.add call $sum end)
        (func (export "_start") i32.const 10 i32.const 0 call $sum call $log))
