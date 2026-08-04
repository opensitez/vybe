;; vybe-test: wast/wat_extended_const/test_const_add_in_global_init
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
        (global $g i32 (i32.add (i32.const 40) (i32.const 2)))
        (func (export "_start") global.get $g i32.const 42 call $vybe_check_i32))
