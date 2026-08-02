;; vybe-test: wast/wat_algorithms/test_integer_sqrt_newton
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $isqrt (param $n i32) (result i32) (local $x i32) (local $y i32)
          local.get $n local.set $x
          local.get $n i32.const 1 i32.add i32.const 2 i32.div_u local.set $y
          block loop
            local.get $y local.get $x i32.lt_u i32.eqz br_if 1
            local.get $y local.set $x
            local.get $y local.get $n local.get $y i32.div_u i32.add i32.const 2 i32.div_u local.set $y
            br 0
          end end local.get $x)
        (func (export "_start") i32.const 144 call $isqrt call $log))
