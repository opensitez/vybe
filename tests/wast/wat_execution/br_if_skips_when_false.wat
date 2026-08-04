;; vybe-test: wast/wat_execution/br_if_skips_when_false
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
    block $b
      i32.const 0
      br_if $b
      i32.const 5
      i32.const 5 call $vybe_check_i32
    end
    i32.const 6
    i32.const 6 call $vybe_check_i32))
