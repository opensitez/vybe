;; vybe-test: wast/wat_execution/multi_return_values_used
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
  (func $divmod (param $a i32) (param $b i32) (result i32 i32)
    local.get $a local.get $b i32.div_u
    local.get $a local.get $b i32.rem_u)
  (func (export "_start")
    i32.const 17
    i32.const 5
    call $divmod
    i32.const 2 call $vybe_check_i32   ;; rem = 2
    i32.const 3 call $vybe_check_i32)) ;; div = 3
