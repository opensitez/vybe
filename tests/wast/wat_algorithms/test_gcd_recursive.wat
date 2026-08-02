;; vybe-test: wast/wat_algorithms/test_gcd_recursive
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $gcd (param $a i32) (param $b i32) (result i32)
          local.get $b i32.eqz
          if (result i32) local.get $a
          else local.get $b local.get $a local.get $b i32.rem_u call $gcd end)
        (func (export "_start") i32.const 1071 i32.const 462 call $gcd call $log))
