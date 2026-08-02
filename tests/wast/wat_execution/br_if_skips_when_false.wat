;; vybe-test: wast/wat_execution/br_if_skips_when_false
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    block $b
      i32.const 0
      br_if $b
      i32.const 5
      call $log
    end
    i32.const 6
    call $log))
