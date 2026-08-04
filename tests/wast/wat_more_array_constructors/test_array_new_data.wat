;; vybe-test: wast/wat_more_array_constructors/test_array_new_data
;; origin: languages/wast/tests/wast/test_wat_more_array_constructors.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (type $a (array (mut i32)))
        (data $d "\63\00\00\00\64\00\00\00")
        (func (export "_start")
          i32.const 0 i32.const 2 array.new_data $a $d i32.const 0 array.get $a i32.const 99 call $vybe_check_i32))
