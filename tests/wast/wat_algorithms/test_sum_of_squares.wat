;; vybe-test: wast/wat_algorithms/test_sum_of_squares
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") (local $i i32) (local $s i32) i32.const 1 local.set $i
          block loop local.get $i i32.const 5 i32.gt_s br_if 1
            local.get $s local.get $i local.get $i i32.mul i32.add local.set $s
            local.get $i i32.const 1 i32.add local.set $i br 0 end end
          local.get $s call $log))
