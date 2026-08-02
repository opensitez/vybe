;; vybe-test: wast/wat_algorithms/test_power_mod
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $powmod (param $b i32) (param $e i32) (param $m i32) (result i32) (local $r i32)
          i32.const 1 local.set $r
          block loop local.get $e i32.eqz br_if 1
            local.get $r local.get $b i32.mul local.get $m i32.rem_u local.set $r
            local.get $e i32.const 1 i32.sub local.set $e br 0 end end
          local.get $r)
        (func (export "_start") i32.const 2 i32.const 10 i32.const 1000 call $powmod call $log))
