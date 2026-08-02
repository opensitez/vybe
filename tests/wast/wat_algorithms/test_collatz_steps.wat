;; vybe-test: wast/wat_algorithms/test_collatz_steps
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $collatz (param $n i32) (result i32) (local $steps i32)
          block loop
            local.get $n i32.const 1 i32.le_u br_if 1
            local.get $n i32.const 1 i32.and
            if local.get $n i32.const 3 i32.mul i32.const 1 i32.add local.set $n
            else local.get $n i32.const 2 i32.div_u local.set $n end
            local.get $steps i32.const 1 i32.add local.set $steps br 0
          end end local.get $steps)
        (func (export "_start") i32.const 6 call $collatz call $log))
