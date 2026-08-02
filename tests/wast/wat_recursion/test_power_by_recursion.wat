;; vybe-test: wast/wat_recursion/test_power_by_recursion
;; origin: languages/wast/tests/wast/test_wat_recursion.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $pow (param $b i32) (param $e i32) (result i32)
          local.get $e i32.eqz
          if (result i32) i32.const 1
          else local.get $b local.get $b local.get $e i32.const 1 i32.sub call $pow i32.mul end)
        (func (export "_start") i32.const 2 i32.const 10 call $pow call $log))
