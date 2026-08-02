;; vybe-test: wast/wat_conditionals/test_max_of_two
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $max (param $a i32) (param $b i32) (result i32)
          local.get $a local.get $b local.get $a local.get $b i32.gt_s select)
        (func (export "_start") i32.const 8 i32.const 3 call $max call $log))
