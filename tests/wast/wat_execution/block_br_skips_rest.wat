;; vybe-test: wast/wat_execution/block_br_skips_rest
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    block $b
      i32.const 7
      call $log
      br $b
      i32.const 99
      call $log
    end))
