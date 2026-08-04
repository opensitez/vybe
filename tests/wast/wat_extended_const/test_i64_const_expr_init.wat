;; vybe-test: wast/wat_extended_const/test_i64_const_expr_init
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
        (global $g i64 (i64.mul (i64.const 1000000) (i64.const 1000000)))
        (func (export "_start") global.get $g i64.const 1000000000000 call $vybe_check_i64))
