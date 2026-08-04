;; vybe-test: wast/wat_execution/folded_nested_if_executed
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
    (i32.const 100 call $vybe_check_i32
      (if (result i32) (i32.const 1)
        (then (i32.const 100))
        (else (i32.const 200))))))
