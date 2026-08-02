;; vybe-test: wast/wat_execution/multi_return_values_used
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $divmod (param $a i32) (param $b i32) (result i32 i32)
    local.get $a local.get $b i32.div_u
    local.get $a local.get $b i32.rem_u)
  (func (export "_start")
    i32.const 17
    i32.const 5
    call $divmod
    call $log   ;; rem = 2
    call $log)) ;; div = 3
