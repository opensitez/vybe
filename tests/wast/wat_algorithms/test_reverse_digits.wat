;; vybe-test: wast/wat_algorithms/test_reverse_digits
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $rev (param $n i32) (result i32) (local $r i32)
          block loop local.get $n i32.eqz br_if 1
            local.get $r i32.const 10 i32.mul local.get $n i32.const 10 i32.rem_u i32.add local.set $r
            local.get $n i32.const 10 i32.div_u local.set $n br 0 end end
          local.get $r)
        (func (export "_start") i32.const 12345 call $rev call $log))
