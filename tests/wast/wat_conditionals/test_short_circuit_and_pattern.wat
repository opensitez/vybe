;; vybe-test: wast/wat_conditionals/test_short_circuit_and_pattern
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func $both (param $a i32) (param $b i32) (result i32)
          local.get $a if (result i32) local.get $b else i32.const 0 end)
        (func (export "_start") i32.const 1 i32.const 5 call $both i32.const 5 call $vybe_check_i32))
