;; vybe-test: wast/wat_execution/recursive_factorial
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $fact (param $n i32) (result i32)
    local.get $n
    i32.const 1
    i32.le_s
    if (result i32)
      i32.const 1
    else
      local.get $n
      local.get $n
      i32.const 1
      i32.sub
      call $fact
      i32.mul
    end)
  (func (export "_start")
    i32.const 6
    call $fact
    call $log))
