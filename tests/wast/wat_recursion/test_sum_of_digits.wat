;; vybe-test: wast/wat_recursion/test_sum_of_digits
;; origin: languages/wast/tests/wast/test_wat_recursion.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $dig (param $n i32) (result i32)
          local.get $n i32.const 10 i32.lt_u
          if (result i32) local.get $n
          else local.get $n i32.const 10 i32.rem_u
               local.get $n i32.const 10 i32.div_u call $dig i32.add end)
        (func (export "_start") i32.const 12345 call $dig call $log))
