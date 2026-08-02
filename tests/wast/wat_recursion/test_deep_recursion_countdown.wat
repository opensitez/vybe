;; vybe-test: wast/wat_recursion/test_deep_recursion_countdown
;; origin: languages/wast/tests/wast/test_wat_recursion.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $count (param $n i32) (result i32)
          local.get $n i32.eqz
          if (result i32) i32.const 0
          else local.get $n i32.const 1 i32.sub call $count i32.const 1 i32.add end)
        (func (export "_start") i32.const 500 call $count call $log))
