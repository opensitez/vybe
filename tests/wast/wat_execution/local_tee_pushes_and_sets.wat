;; vybe-test: wast/wat_execution/local_tee_pushes_and_sets
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $x i32)
    i32.const 55
    local.tee $x
    call $log))
