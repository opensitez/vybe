;; vybe-test: wast/wat_extended_const/test_nested_const_expr
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (global $g i32 (i32.add (i32.mul (i32.const 5) (i32.const 8)) (i32.const 2)))
        (func (export "_start") global.get $g i32.const 42 call $vybe_check_i32))
