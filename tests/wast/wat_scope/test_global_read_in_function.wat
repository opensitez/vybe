;; vybe-test: wast/wat_scope/test_global_read_in_function
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (global $base i32 (i32.const 1000))
        (func $offset (param $d i32) (result i32) global.get $base local.get $d i32.add)
        (func (export "_start") i32.const 23 call $offset i32.const 1023 call $vybe_check_i32))
