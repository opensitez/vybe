;; vybe-test: wast/wat_conditionals/test_even_or_odd
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
        (func $parity (param $n i32) (result i32) local.get $n i32.const 1 i32.and)
        (func (export "_start") i32.const 13 call $parity i32.const 1 call $vybe_check_i32))
