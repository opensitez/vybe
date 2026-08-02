;; vybe-test: wast/wat_conditionals/test_short_circuit_and_pattern
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $both (param $a i32) (param $b i32) (result i32)
          local.get $a if (result i32) local.get $b else i32.const 0 end)
        (func (export "_start") i32.const 1 i32.const 5 call $both call $log))
